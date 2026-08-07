# -*- coding: utf-8 -*-
"""
Outlook メール変換ツール

Outlook からフォルダーへドラッグ＆ドロップ保存された .eml / .msg を検知し、
「その場所」で本文（HTML / TXT）と添付ファイルに分離して保存する。
監視範囲は config.ini の watch_roots（既定: ドキュメントとデスクトップ配下すべて）。

動作モード（config.ini の mode）:
  auto = 設定に従って黙って変換する
  ask  = ドロップのたびに選択画面を出す
         （フォルダ作成有無・フォルダー名・添付保存有無・本文ファイル名・形式）

保存スタイル（config.ini の layout / 選択画面のチェック）:
  folder = 「YYYYMMDD_件名」フォルダーを作ってその中に保存
  flat   = フォルダーを作らず、ドロップした場所に直接保存

使い方:
  python mail_converter.py --watch           常駐監視（新着を数秒で自動変換）
  python mail_converter.py --once            巡回対象フォルダーを一括変換して終了
  python mail_converter.py --convert <フォルダー>  指定フォルダー内をまとめて変換
  python mail_converter.py <ファイル>         指定した 1 ファイルだけ変換

設定は同じフォルダーの config.ini を編集する（変更後は start_watcher.cmd で再起動）。
"""

import base64
import configparser
import ctypes
import email
import html as html_mod
import json
import mimetypes
import os
import queue as queue_mod
import re
import socket
import sys
import time
import traceback
from email import policy
from email.utils import parsedate_to_datetime
from pathlib import Path

TOOL_DIR = Path(__file__).resolve().parent
CONFIG_PATH = TOOL_DIR / "config.ini"
STATE_PATH = TOOL_DIR / "state.json"
LOG_PATH = TOOL_DIR / "converter.log"
LOCK_PORT = 53211  # 常駐の二重起動防止用（ローカル専用ポート）

INVALID_CHARS = re.compile(r'[\\/:*?"<>|\r\n\t]')

# この名前を含む場所は変換対象にしない（開発リポジトリ等の巻き込み防止）
EXCLUDE_DIR_NAMES = {"node_modules", ".git", ".venv", "__pycache__", "AppData"}

# 危険な HTML 部品の除去パターン（ローカルで開いても安全なように）
SCRIPT_RE = re.compile(r"<script\b.*?</script\s*>", re.I | re.S)
IFRAME_RE = re.compile(r"<iframe\b.*?(?:</iframe\s*>|/\s*>)", re.I | re.S)
OBJECT_RE = re.compile(r"<object\b.*?</object\s*>", re.I | re.S)
EMBED_RE = re.compile(r"<embed\b[^>]*>", re.I)
META_REFRESH_RE = re.compile(r"<meta\b[^>]*http-equiv\s*=\s*[\"']?refresh[^>]*>", re.I)
ON_ATTR_RE = re.compile(r"\son\w+\s*=\s*(?:\"[^\"]*\"|'[^']*'|[^\s>]+)", re.I)
JS_URL_RE = re.compile(r"(href|src)(\s*=\s*[\"']?)\s*javascript:", re.I)
REMOTE_IMG_RE = re.compile(r"(<img\b[^>]*?)\bsrc(\s*=\s*[\"']https?:)", re.I)


class Unconvertible(Exception):
    """本文を取り出せないメール（暗号化など）。元ファイルを壊さず残す。"""


class UserSkipped(Exception):
    """選択画面で「変換しない」が選ばれた。元ファイルはそのまま残す。"""


class MailContent:
    def __init__(self):
        self.subject = "(件名なし)"
        self.sender = ""
        self.to = ""
        self.date_str = ""
        self.stamp = None       # フォルダー/ファイル名用 YYYYMMDD
        self.html = None
        self.text = None
        self.inline = []        # (cid, ファイル名, バイト列)
        self.attachments = []   # (ファイル名, バイト列)
        self.encrypted = False


# ---------------------------------------------------------------- 共通部品

def sanitize(name, maxlen=60):
    name = INVALID_CHARS.sub("_", str(name)).strip().rstrip(".")
    return name[:maxlen].rstrip() or "no_name"


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    for i in range(2, 1000):
        cand = path.with_name(f"{path.stem}_{i}{path.suffix}")
        if not cand.exists():
            return cand
    raise RuntimeError(f"同名ファイルが多すぎます: {path}")


def sanitize_html(body: str, remote_images: bool) -> str:
    body = SCRIPT_RE.sub("", body)
    body = IFRAME_RE.sub("", body)
    body = OBJECT_RE.sub("", body)
    body = EMBED_RE.sub("", body)
    body = META_REFRESH_RE.sub("", body)
    body = ON_ATTR_RE.sub("", body)
    body = JS_URL_RE.sub(r"\1\2blocked:", body)
    if not remote_images:
        body = REMOTE_IMG_RE.sub(r"\1data-blocked-src\2", body)
    return body


def html_to_text(body: str) -> str:
    body = re.sub(r"(?is)<(script|style)\b.*?</\1\s*>", "", body)
    body = re.sub(r"(?i)<br\s*/?>", "\n", body)
    body = re.sub(r"(?i)</(p|div|tr|li|h[1-6]|table)\s*>", "\n", body)
    body = re.sub(r"<[^>]+>", "", body)
    body = html_mod.unescape(body)
    body = re.sub(r"[ \t]+\n", "\n", body)
    body = re.sub(r"\n{3,}", "\n\n", body)
    return body.strip()


def part_text(part) -> str:
    """charset 不備のメールでも落ちないテキスト取り出し。"""
    try:
        return part.get_content()
    except Exception:
        raw = part.get_payload(decode=True) or b""
        cs = part.get_content_charset() or "utf-8"
        try:
            return raw.decode(cs, errors="replace")
        except LookupError:
            return raw.decode("utf-8", errors="replace")


def norm_dir(p) -> str:
    """フォルダーパスの比較用正規化（大文字小文字と末尾の \\ を無視）。"""
    return str(Path(p).resolve()).rstrip("\\/").lower()


# ---------------------------------------------------------------- .eml 解析

def parse_eml(path: Path) -> MailContent:
    with open(path, "rb") as f:
        msg = email.message_from_binary_file(f, policy=policy.default)
    c = MailContent()
    c.subject = str(msg.get("Subject", "") or "(件名なし)")
    c.sender = str(msg.get("From", "") or "")
    c.to = str(msg.get("To", "") or "")
    c.date_str = str(msg.get("Date", "") or "")
    try:
        c.stamp = parsedate_to_datetime(msg["Date"]).astimezone().strftime("%Y%m%d")
    except Exception:
        c.stamp = None
    _walk_eml(msg, c)
    # 本文 HTML から cid: で参照されている画像だけをインライン扱いにし、
    # 参照されていないものは通常の添付として保存する
    referenced, others = [], []
    for cid, name, data in c.inline:
        if c.html and f"cid:{cid}" in c.html:
            referenced.append((cid, name, data))
        else:
            others.append((name, data))
    c.inline = referenced
    c.attachments.extend(others)
    return c


def _walk_eml(part, c: MailContent):
    ct = part.get_content_type()
    if ct == "message/rfc822":
        # 転送メールの入れ子はそのまま .eml として添付扱いで保存
        try:
            inner = part.get_payload(0)
            name = sanitize(str(inner.get("Subject", "") or "添付メール"), 50) + ".eml"
            c.attachments.append((name, inner.as_bytes()))
        except Exception:
            c.attachments.append(("添付メール.eml", part.get_payload(decode=True) or b""))
        return
    if part.is_multipart():
        for sub in part.iter_parts():
            _walk_eml(sub, c)
        return

    cd = part.get_content_disposition()
    fn = part.get_filename()
    cid = (part.get("Content-ID") or "").strip("<>")

    if ct in ("application/pkcs7-signature", "application/x-pkcs7-signature"):
        return  # 電子署名データは保存対象外
    if ct == "application/pkcs7-mime":
        c.encrypted = True
        return
    if ct == "text/html" and cd != "attachment" and c.html is None:
        c.html = part_text(part)
    elif ct == "text/plain" and cd != "attachment" and c.text is None:
        c.text = part_text(part)
    elif ct == "text/calendar":
        data = part.get_payload(decode=True) or part_text(part).encode("utf-8")
        c.attachments.append((fn or "invite.ics", data))
    elif cid and ct.startswith("image/"):
        # インライン画像候補。本文から参照されているかは parse_eml で最終判定する
        ext = ct.split("/")[1].split("+")[0]
        c.inline.append((cid, fn or f"image.{ext}", part.get_payload(decode=True) or b""))
    elif fn or cd == "attachment":
        c.attachments.append((fn or "attachment.bin", part.get_payload(decode=True) or b""))


# ---------------------------------------------------------------- .msg 解析

def parse_msg(path: Path) -> MailContent:
    import extract_msg  # classic Outlook の .msg 形式用

    c = MailContent()
    msg = extract_msg.openMsg(str(path))
    if not getattr(msg, "areStringsUnicode", True):
        # 旧式の ANSI 形式 .msg は日本語環境の文字コード (cp932) で読み直す
        try:
            msg.close()
            msg = extract_msg.openMsg(str(path), overrideEncoding="cp932")
        except Exception:
            msg = extract_msg.openMsg(str(path))
    try:
        c.subject = msg.subject or "(件名なし)"
        c.sender = msg.sender or ""
        c.to = msg.to or ""
        d = msg.date
        if d is not None:
            try:
                c.stamp = d.astimezone().strftime("%Y%m%d")
                c.date_str = d.astimezone().strftime("%Y-%m-%d %H:%M")
            except Exception:
                c.date_str = str(d)
        hb = msg.htmlBody
        if hb:
            c.html = hb.decode("utf-8", errors="replace") if isinstance(hb, bytes) else hb
        if msg.body:
            c.text = msg.body
        for att in msg.attachments:
            name = att.longFilename or att.shortFilename or "attachment.bin"
            data = att.data
            if isinstance(data, bytes):
                cid = getattr(att, "cid", None)
                if cid and c.html and f"cid:{cid}" in c.html:
                    c.inline.append((cid, name, data))
                else:
                    c.attachments.append((name, data))
            else:
                # 埋め込みメッセージ（転送メール等）は .msg のまま保存
                try:
                    raw = data.exportBytes()
                    if not name.lower().endswith(".msg"):
                        name += ".msg"
                    c.attachments.append((name, raw))
                except Exception:
                    pass
    finally:
        msg.close()
    return c


# ---------------------------------------------------------------- 変換オプション

def resolve_layout(target_dir: Path, cfg) -> str:
    d = norm_dir(target_dir)
    if d in cfg["flat_folders"]:
        return "flat"
    if d in cfg["folder_folders"]:
        return "folder"
    return cfg["layout"]


def build_default_opts(c: MailContent, target_dir: Path, cfg) -> dict:
    layout = resolve_layout(target_dir, cfg)
    base = f"{c.stamp or '日付不明'}_{sanitize(c.subject, 60)}"
    return {
        "layout": layout,
        "dir_name": base,
        "body_base": "mail" if layout == "folder" else base,
        "save_attachments": True,
        "format": cfg["format"],
    }


# Machi-AI デザイン基盤 v3・夜面トークン（選択画面用）
DIALOG_C = {
    "bg": "#14171C",        # 面
    "bg2": "#181C22",       # 面2（入力欄・チップ）
    "bg3": "#1E232A",       # 面3（ホバー）
    "line": "#2A2C33",      # 線
    "line2": "#3C3F48",     # 強線
    "text": "#E8EBEF",      # 文字（純白禁止）
    "text2": "#B4BCC7",     # 文字2
    "text3": "#7C8794",     # 文字3（静音）
    "mute": "#5A636E",      # 文字4（無効）
    "accent": "#FF9F5A",    # 橙の灯り（1 画面 1 箇所＝主ボタンのみ）
    "accent_hi": "#FFB27A",
    "accent_fg": "#1C1006",  # 橙の上は濃色文字（白文字禁止）
}


def _pretty_date(c: MailContent) -> str:
    try:
        return parsedate_to_datetime(c.date_str).astimezone().strftime("%Y年%m月%d日 %H:%M")
    except Exception:
        return c.date_str or "日時不明"


def ask_user_options(path: Path, c: MailContent, defaults: dict, auto_close_ms=None):
    """ドロップごとの選択画面。戻り値: 変換オプション dict / None（変換しない）

    customtkinter（角丸・スイッチ等のモダン部品）で表示し、
    未導入の環境では従来の tkinter 版に自動フォールバックする。
    """
    try:
        import customtkinter as ctk
    except ImportError:
        return _ask_user_options_tk(path, c, defaults, auto_close_ms)
    import tkinter as tk
    import tkinter.font as tkfont

    C = DIALOG_C
    base = defaults["dir_name"]
    result = {"value": None}
    FMT_LABEL = {"html": "HTML", "txt": "テキスト", "both": "両方"}
    FMT_VALUE = {v: k for k, v in FMT_LABEL.items()}

    ctk.set_appearance_mode("dark")
    root = ctk.CTk(fg_color=C["bg"])
    root.title("メールの保存")
    root.attributes("-topmost", True)
    root.resizable(False, False)
    try:
        # Windows のタイトルバーもダークにする
        root.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
        dark = ctypes.c_int(1)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(dark), 4)
    except Exception:
        pass
    def _set_brand_icon():
        # タイトルバーのアイコンをブランドの「灯り」（墨地に橙の点）へ差し替える。
        # customtkinter が起動直後に .ico を設定するため、後から .ico で上書きする
        # （Windows のタイトルバーは iconphoto より iconbitmap の .ico が優先）
        try:
            ico = TOOL_DIR / "lantern.ico"
            if not ico.exists():
                from PIL import Image, ImageDraw
                img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
                d = ImageDraw.Draw(img)
                d.rounded_rectangle((2, 2, 62, 62), radius=14, fill=C["bg3"])
                d.ellipse((22, 22, 42, 42), fill=C["accent"])
                img.save(ico, sizes=[(16, 16), (32, 32), (48, 48)])
            root.iconbitmap(str(ico))
        except Exception:
            pass

    root.after(400, _set_brand_icon)

    fams = set(tkfont.families())
    fam = next((f for f in ("Zen Kaku Gothic New", "Yu Gothic UI", "Meiryo UI")
                if f in fams), None)

    def font(size, weight="normal"):
        if fam:
            return ctk.CTkFont(family=fam, size=size, weight=weight)
        return ctk.CTkFont(size=size, weight=weight)

    f_kick = font(12)
    f_head = font(17, "bold")
    f_meta = font(12)
    f_body = font(13)
    f_btn = font(13, "bold")

    WRAP = 560  # 本文幅（ここが実質のウィンドウ幅になる）
    frm = ctk.CTkFrame(root, fg_color=C["bg"], corner_radius=0)
    frm.pack(fill="both", expand=True, padx=28, pady=24)

    def lab(text, fnt, color, **pack_opts):
        w = ctk.CTkLabel(frm, text=text, font=fnt, text_color=color,
                         anchor="w", justify="left")
        w.pack(fill="x", **pack_opts)
        return w

    def divider(pady):
        ctk.CTkFrame(frm, fg_color=C["line"], height=1,
                     corner_radius=0).pack(fill="x", pady=pady)

    # ---- 見出し部
    lab(f"保存先　{path.parent.name}", f_kick, C["text3"])
    subj = c.subject if len(c.subject) <= 80 else c.subject[:80] + "…"
    w_subj = lab(subj, f_head, C["text"], pady=(6, 2))
    w_subj.configure(wraplength=WRAP)
    lab(f"{_pretty_date(c)}　・　添付 {len(c.attachments)} 件",
        f_meta, C["text2"])
    divider((14, 18))

    # ---- 入力部
    var_folder = tk.BooleanVar(value=defaults["layout"] == "folder")
    var_att = tk.BooleanVar(value=bool(c.attachments))
    var_dir = tk.StringVar(value=base)
    var_body = tk.StringVar(value="mail" if defaults["layout"] == "folder" else base)
    body_edited = {"flag": False}

    def on_body_key(_e):
        body_edited["flag"] = True

    def on_toggle_folder():
        on_ = var_folder.get()
        ent_dir.configure(state="normal" if on_ else "disabled")
        w_dir_lab.configure(text_color=C["text2"] if on_ else C["mute"])
        if not body_edited["flag"]:
            var_body.set("mail" if on_ else base)

    def mk_switch(text, var, cmd=None):
        sw = ctk.CTkSwitch(
            frm, text=text, variable=var, command=cmd, font=f_body,
            text_color=C["text"], text_color_disabled=C["mute"],
            fg_color=C["line"], progress_color=C["line2"],
            button_color=C["text3"], button_hover_color=C["text"],
            switch_width=44, switch_height=22)

        def _sync(*_a):
            # ON はつまみを明るく灯し、OFF は暗く沈める（位置＋明度の二重サイン）
            sw.configure(button_color=C["text"] if var.get() else C["text3"])

        var.trace_add("write", _sync)
        _sync()
        return sw

    def mk_entry(var):
        # width がウィンドウ全体の横幅を決める（入力しやすさ優先で広め）
        return ctk.CTkEntry(
            frm, textvariable=var, font=f_body, height=38, width=WRAP,
            fg_color=C["bg2"], border_color=C["line"], border_width=1,
            corner_radius=8, text_color=C["text"],
            placeholder_text_color=C["mute"])

    mk_switch("フォルダーを作成して保存する", var_folder,
              on_toggle_folder).pack(anchor="w")
    w_dir_lab = lab("フォルダー名", f_meta, C["text2"], pady=(12, 4))
    ent_dir = mk_entry(var_dir)
    ent_dir.pack(fill="x")

    sw_att = mk_switch(f"添付ファイルを保存する（{len(c.attachments)} 件）", var_att)
    sw_att.pack(anchor="w", pady=(20, 0))
    if not c.attachments:
        sw_att.configure(state="disabled")

    lab("本文ファイル名", f_meta, C["text2"], pady=(20, 4))
    ent_body = mk_entry(var_body)
    ent_body.pack(fill="x")
    ent_body.bind("<Key>", on_body_key)

    lab("本文の形式", f_meta, C["text2"], pady=(20, 6))
    seg = ctk.CTkSegmentedButton(
        frm, values=list(FMT_LABEL.values()), font=f_body, height=36,
        corner_radius=8, border_width=1,
        fg_color=C["line"], selected_color=C["line2"],
        selected_hover_color=C["line2"], unselected_color=C["bg2"],
        unselected_hover_color=C["bg3"], text_color=C["text"])
    seg.set(FMT_LABEL.get(defaults["format"], "HTML"))
    seg.pack(anchor="w")

    divider((22, 16))

    # ---- ボタン部（橙はここ 1 箇所だけ）
    def on_convert(_e=None):
        result["value"] = {
            "layout": "folder" if var_folder.get() else "flat",
            "dir_name": sanitize(var_dir.get(), 60) if var_folder.get() else "",
            "body_base": sanitize(var_body.get(), 60),
            "save_attachments": bool(var_att.get() and c.attachments),
            "format": FMT_VALUE.get(seg.get(), "html"),
        }
        root.destroy()

    def on_skip(_e=None):
        result["value"] = None
        root.destroy()

    btns = ctk.CTkFrame(frm, fg_color=C["bg"], corner_radius=0)
    btns.pack(fill="x")
    ctk.CTkButton(
        btns, text="変換する", command=on_convert, font=f_btn,
        fg_color=C["accent"], hover_color=C["accent_hi"],
        text_color=C["accent_fg"], corner_radius=8, height=40,
        width=150).pack(side="right")
    ctk.CTkButton(
        btns, text="変換しない", command=on_skip, font=f_body,
        fg_color="transparent", hover_color=C["bg3"], text_color=C["text2"],
        border_width=1, border_color=C["line2"], corner_radius=8, height=40,
        width=120).pack(side="right", padx=(0, 12))

    root.bind("<Return>", on_convert)
    root.bind("<Escape>", on_skip)
    root.protocol("WM_DELETE_WINDOW", on_skip)
    on_toggle_folder()
    root.update_idletasks()
    # 画面中央よりやや上に表示して前面へ
    x = (root.winfo_screenwidth() - root.winfo_width()) // 2
    root.geometry(f"+{x}+{root.winfo_screenheight() // 4}")
    root.lift()
    root.focus_force()
    if auto_close_ms:
        root.after(auto_close_ms, on_skip)
    root.mainloop()
    return result["value"]


def _ask_user_options_tk(path: Path, c: MailContent, defaults: dict, auto_close_ms=None):
    """従来の tkinter 版（customtkinter 未導入環境の予備）。"""
    import tkinter as tk
    import tkinter.font as tkfont

    C = DIALOG_C
    base = defaults["dir_name"]
    result = {"value": None}

    root = tk.Tk()
    root.title("メールの保存")
    root.configure(bg=C["bg"])
    root.attributes("-topmost", True)
    root.resizable(False, False)
    try:
        # Windows のタイトルバーもダークにする
        root.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
        dark = ctypes.c_int(1)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(dark), 4)
    except Exception:
        pass

    fams = set(tkfont.families())
    fam = next((f for f in ("Zen Kaku Gothic New", "Yu Gothic UI", "Meiryo UI")
                if f in fams), "TkDefaultFont")
    f_label = (fam, 9, "bold")
    f_head = (fam, 12, "bold")
    f_meta = (fam, 9)
    f_body = (fam, 10)
    f_btn = (fam, 10, "bold")

    frm = tk.Frame(root, bg=C["bg"], padx=24, pady=20)
    frm.pack(fill="both", expand=True)
    frm.grid_columnconfigure(1, weight=1)

    def lab(text, font, fg, **grid):
        w = tk.Label(frm, text=text, font=font, bg=C["bg"], fg=fg,
                     anchor="w", justify="left")
        w.grid(**grid)
        return w

    def divider(row, pady):
        tk.Frame(frm, bg=C["line"], height=1).grid(
            row=row, column=0, columnspan=2, sticky="ew", pady=pady)

    # ---- 見出し部
    subj = c.subject if len(c.subject) <= 60 else c.subject[:60] + "…"
    w_subj = lab(subj, f_head, C["text"],
                 row=1, column=0, columnspan=2, sticky="w", pady=(0, 4))
    w_subj.configure(wraplength=440)
    lab(f"{_pretty_date(c)}　・　添付 {len(c.attachments)} 件", f_meta, C["text2"],
        row=2, column=0, columnspan=2, sticky="w", pady=(0, 2))
    lab(f"保存先　{path.parent.name}", f_meta, C["text3"],
        row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
    divider(3, (12, 16))

    # ---- 入力部
    var_folder = tk.BooleanVar(value=defaults["layout"] == "folder")
    var_att = tk.BooleanVar(value=bool(c.attachments))
    var_fmt = tk.StringVar(value=defaults["format"])
    var_dir = tk.StringVar(value=base)
    var_body = tk.StringVar(value="mail" if defaults["layout"] == "folder" else base)
    body_edited = {"flag": False}

    def mk_check(text, var, cmd=None):
        return tk.Checkbutton(
            frm, text=text, variable=var, command=cmd, font=f_body,
            bg=C["bg"], fg=C["text"], activebackground=C["bg"],
            activeforeground=C["text"], selectcolor=C["bg2"],
            disabledforeground=C["mute"], highlightthickness=0, bd=0, anchor="w")

    def mk_entry(var):
        return tk.Entry(
            frm, textvariable=var, font=f_body, bg=C["bg2"], fg=C["text"],
            insertbackground=C["text"], relief="flat", highlightthickness=1,
            highlightbackground=C["line"], highlightcolor=C["line2"],
            disabledbackground=C["bg"], disabledforeground=C["mute"])

    def on_body_key(_e):
        body_edited["flag"] = True

    def on_toggle_folder():
        on_ = var_folder.get()
        ent_dir.config(state="normal" if on_ else "disabled")
        w_dir_lab.config(fg=C["text2"] if on_ else C["mute"])
        if not body_edited["flag"]:
            var_body.set("mail" if on_ else base)

    mk_check("フォルダーを作成して保存する", var_folder, on_toggle_folder).grid(
        row=4, column=0, columnspan=2, sticky="w")
    w_dir_lab = lab("フォルダー名", f_meta, C["text2"],
                    row=5, column=0, columnspan=2, sticky="w",
                    padx=(28, 0), pady=(6, 2))
    ent_dir = mk_entry(var_dir)
    ent_dir.grid(row=6, column=0, columnspan=2, sticky="ew",
                 padx=(28, 0), ipady=5)

    cb_att = mk_check(f"添付ファイルを保存する（{len(c.attachments)} 件）", var_att)
    cb_att.grid(row=7, column=0, columnspan=2, sticky="w", pady=(14, 0))
    if not c.attachments:
        cb_att.config(state="disabled")

    lab("本文ファイル名", f_meta, C["text2"],
        row=8, column=0, columnspan=2, sticky="w", padx=(28, 0), pady=(14, 2))
    ent_body = mk_entry(var_body)
    ent_body.grid(row=9, column=0, columnspan=2, sticky="ew",
                  padx=(28, 0), ipady=5)
    ent_body.bind("<Key>", on_body_key)

    lab("本文の形式", f_meta, C["text2"],
        row=10, column=0, columnspan=2, sticky="w", padx=(28, 0), pady=(14, 2))
    seg = tk.Frame(frm, bg=C["bg"])
    seg.grid(row=11, column=0, columnspan=2, sticky="w", padx=(28, 0))
    f_seg_on = (fam, 10, "bold")
    seg_btns = []
    for text, val in (("HTML", "html"), ("テキスト", "txt"), ("両方", "both")):
        rb = tk.Radiobutton(
            seg, text=text, value=val, variable=var_fmt, indicatoron=0,
            font=f_body, bg=C["bg2"], fg=C["text3"], selectcolor=C["line2"],
            activebackground=C["bg3"], activeforeground=C["text"],
            relief="flat", offrelief="flat", bd=0, padx=16, pady=5,
            highlightthickness=0, cursor="hand2")
        rb.pack(side="left", padx=(0, 6))
        seg_btns.append((rb, val))

    def on_fmt_change(*_a):
        # 選択中の形式だけ太字・明色にして識別しやすくする
        for rb, val in seg_btns:
            if var_fmt.get() == val:
                rb.config(font=f_seg_on, fg=C["text"])
            else:
                rb.config(font=f_body, fg=C["text3"])

    var_fmt.trace_add("write", on_fmt_change)
    on_fmt_change()

    divider(12, (18, 14))

    # ---- ボタン部（橙はここ 1 箇所だけ）
    def on_convert(_e=None):
        result["value"] = {
            "layout": "folder" if var_folder.get() else "flat",
            "dir_name": sanitize(var_dir.get(), 60) if var_folder.get() else "",
            "body_base": sanitize(var_body.get(), 60),
            "save_attachments": bool(var_att.get() and c.attachments),
            "format": var_fmt.get(),
        }
        root.destroy()

    def on_skip(_e=None):
        result["value"] = None
        root.destroy()

    btns = tk.Frame(frm, bg=C["bg"])
    btns.grid(row=13, column=0, columnspan=2, sticky="e")
    ghost = tk.Button(
        btns, text="変換しない", command=on_skip, font=f_body,
        bg=C["bg"], fg=C["text2"], activebackground=C["bg3"],
        activeforeground=C["text"], relief="flat", bd=0, padx=18, pady=7,
        highlightthickness=1, highlightbackground=C["line2"], cursor="hand2")
    ghost.pack(side="left", padx=(0, 10))
    primary = tk.Button(
        btns, text="変換する", command=on_convert, font=f_btn,
        bg=C["accent"], fg=C["accent_fg"], activebackground=C["accent_hi"],
        activeforeground=C["accent_fg"], relief="flat", bd=0, padx=24, pady=7,
        highlightthickness=0, cursor="hand2")
    primary.pack(side="left")
    primary.bind("<Enter>", lambda e: primary.config(bg=C["accent_hi"]))
    primary.bind("<Leave>", lambda e: primary.config(bg=C["accent"]))
    ghost.bind("<Enter>", lambda e: ghost.config(bg=C["bg3"]))
    ghost.bind("<Leave>", lambda e: ghost.config(bg=C["bg"]))

    root.bind("<Return>", on_convert)
    root.bind("<Escape>", on_skip)
    root.protocol("WM_DELETE_WINDOW", on_skip)
    on_toggle_folder()
    root.update_idletasks()
    # 画面中央よりやや上に表示して前面へ
    x = (root.winfo_screenwidth() - root.winfo_width()) // 2
    root.geometry(f"+{x}+{root.winfo_screenheight() // 4}")
    root.lift()
    root.focus_force()
    if auto_close_ms:
        root.after(auto_close_ms, on_skip)
    root.mainloop()
    return result["value"]


# ---------------------------------------------------------------- 出力

def _header_html(c: MailContent, att_html: str) -> str:
    esc = html_mod.escape
    return (
        '<div style="border-bottom:2px solid #999;padding:10px;margin-bottom:10px;'
        'font-family:sans-serif;font-size:13px;background:#f7f7f7;color:#222">'
        f"<b>件名:</b> {esc(c.subject)}<br>"
        f"<b>差出人:</b> {esc(c.sender)}<br>"
        f"<b>宛先:</b> {esc(c.to)}<br>"
        f"<b>日時:</b> {esc(c.date_str)}"
        + (f"<br><b>添付:</b> {att_html}" if att_html else "")
        + "</div>\n"
    )


def _header_txt(c: MailContent, att_line: str) -> str:
    return (
        f"件名: {c.subject}\n差出人: {c.sender}\n宛先: {c.to}\n日時: {c.date_str}\n"
        f"添付: {att_line or 'なし'}\n" + "-" * 40 + "\n"
    )


def _html_page(c: MailContent, cid_map: dict, header: str, cfg) -> str:
    esc = html_mod.escape
    if c.html is not None:
        body = c.html
        for cid, rel in cid_map.items():
            body = body.replace(f"cid:{cid}", rel)
        body = sanitize_html(body, cfg["remote_images"])
        return f'<meta charset="utf-8">\n{header}{body}'
    text = c.text or "(本文なし)"
    return (f'<meta charset="utf-8">\n{header}'
            f'<pre style="white-space:pre-wrap;font-family:sans-serif">{esc(text)}</pre>')


def write_output(c: MailContent, target_dir: Path, cfg, opts: dict) -> dict:
    """opts（layout / dir_name / body_base / save_attachments / format）に従って書き出す。

    folder: dir_name フォルダー + 本文 + attachments/ + assets/
    flat  : 本文ファイルを直接置き、添付は元の名前のまま隣に置く
            （インライン画像は HTML 内に埋め込み、フォルダーを一切作らない）
    """
    esc = html_mod.escape
    fmt = opts["format"]
    body_base = opts["body_base"] or "mail"
    saved, att_names, generated = [], [], []

    if opts["layout"] == "folder":
        out_dir = unique_path(target_dir / (opts["dir_name"] or "no_name"))
        out_dir.mkdir(parents=True)

        cid_map = {}
        if c.inline:
            assets = out_dir / "assets"
            assets.mkdir()
            for cid, name, data in c.inline:
                p = unique_path(assets / sanitize(name, 80))
                p.write_bytes(data)
                cid_map[cid] = f"assets/{p.name}"
        if c.attachments and opts["save_attachments"]:
            att_dir = out_dir / "attachments"
            att_dir.mkdir()
            for name, data in c.attachments:
                p = unique_path(att_dir / sanitize(name, 80))
                p.write_bytes(data)
                att_names.append(p.name)
                if p.suffix.lower() in (".eml", ".msg"):
                    generated.append(p)
        att_html = "、".join(
            f'<a href="attachments/{esc(n)}">{esc(n)}</a>' for n in att_names
        )
        att_line = "、".join(att_names)
        if c.attachments and not opts["save_attachments"]:
            names = "、".join(sanitize(n, 80) for n, _ in c.attachments)
            att_html = esc(names) + "（保存しない設定のため未保存）"
            att_line = names + "（未保存）"
        if fmt in ("html", "both"):
            p = unique_path(out_dir / f"{body_base}.html")
            p.write_text(_html_page(c, cid_map, _header_html(c, att_html), cfg),
                         encoding="utf-8")
            saved.append(p.name)
        if fmt in ("txt", "both"):
            body = c.text if c.text is not None else html_to_text(c.html or "(本文なし)")
            p = unique_path(out_dir / f"{body_base}.txt")
            p.write_text(_header_txt(c, att_line) + body + "\n", encoding="utf-8")
            saved.append(p.name)
    else:  # flat
        out_dir = target_dir
        # インライン画像はファイルを作らず HTML へ直接埋め込む
        cid_map = {}
        for cid, name, data in c.inline:
            mime = mimetypes.guess_type(name)[0] or "image/png"
            cid_map[cid] = f"data:{mime};base64,{base64.b64encode(data).decode()}"
        if c.attachments and opts["save_attachments"]:
            for name, data in c.attachments:
                p = unique_path(target_dir / sanitize(name, 80))
                p.write_bytes(data)
                att_names.append(p.name)
                if p.suffix.lower() in (".eml", ".msg"):
                    generated.append(p)
        att_html = "、".join(f'<a href="{esc(n)}">{esc(n)}</a>' for n in att_names)
        att_line = "、".join(att_names)
        if c.attachments and not opts["save_attachments"]:
            names = "、".join(sanitize(n, 80) for n, _ in c.attachments)
            att_html = esc(names) + "（保存しない設定のため未保存）"
            att_line = names + "（未保存）"
        if fmt in ("html", "both"):
            p = unique_path(target_dir / f"{body_base}.html")
            p.write_text(_html_page(c, cid_map, _header_html(c, att_html), cfg),
                         encoding="utf-8")
            saved.append(p.name)
        if fmt in ("txt", "both"):
            body = c.text if c.text is not None else html_to_text(c.html or "(本文なし)")
            p = unique_path(target_dir / f"{body_base}.txt")
            p.write_text(_header_txt(c, att_line) + body + "\n", encoding="utf-8")
            saved.append(p.name)

    return {"out_dir": out_dir, "layout": opts["layout"], "saved": saved,
            "attachments": att_names, "generated": generated}


# ---------------------------------------------------------------- 変換の入口

def parse_mail(path: Path) -> MailContent:
    suffix = path.suffix.lower()
    if suffix == ".eml":
        c = parse_eml(path)
    elif suffix == ".msg":
        c = parse_msg(path)
    else:
        raise Unconvertible(f"対応していない形式です: {suffix}")
    if c.encrypted:
        raise Unconvertible("暗号化メールのため本文を取り出せません")
    if c.html is None and c.text is None and not c.attachments and not c.inline:
        raise Unconvertible("本文も添付も見つかりません")
    return c


def convert_file(path: Path, cfg, opts=None, interactive=False) -> dict:
    c = parse_mail(path)
    if opts is None:
        opts = build_default_opts(c, path.parent, cfg)
        if interactive and cfg.get("mode", "auto") == "ask":
            opts = ask_user_options(path, c, opts)
            if opts is None:
                raise UserSkipped()
    return write_output(c, path.parent, cfg, opts)


def dispose_original(path: Path, cfg, log):
    action = cfg["original"]
    if action == "delete":
        path.unlink()
        log(f"元ファイルを削除: {path.name}")
    elif action == "recycle":
        _send_to_recycle_bin(path)
        log(f"元ファイルをごみ箱へ移動: {path.name}")
    # keep のときは何もしない


def _send_to_recycle_bin(path: Path):
    """Windows のごみ箱へ移動（復元可能な削除）。"""

    class SHFILEOPSTRUCTW(ctypes.Structure):
        _fields_ = [
            ("hwnd", ctypes.c_void_p),
            ("wFunc", ctypes.c_uint),
            ("pFrom", ctypes.c_wchar_p),
            ("pTo", ctypes.c_wchar_p),
            ("fFlags", ctypes.c_uint),
            ("fAnyOperationsAborted", ctypes.c_int),
            ("hNameMappings", ctypes.c_void_p),
            ("lpszProgressTitle", ctypes.c_wchar_p),
        ]

    FO_DELETE = 3
    FOF_ALLOWUNDO = 0x40
    FOF_NO_UI = 0x04 | 0x10 | 0x400 | 0x200
    op = SHFILEOPSTRUCTW(
        hwnd=None, wFunc=FO_DELETE, pFrom=str(path) + "\0", pTo=None,
        fFlags=FOF_ALLOWUNDO | FOF_NO_UI,
        fAnyOperationsAborted=0, hNameMappings=None, lpszProgressTitle=None,
    )
    res = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
    if res != 0:
        raise OSError(f"ごみ箱への移動に失敗しました (code={res}): {path}")


# ---------------------------------------------------------------- 設定・状態・ログ

def _split_paths(value: str):
    return [os.path.expandvars(s.strip()) for s in str(value).split(";") if s.strip()]


def load_config():
    # interpolation=None: 値の中の % を特別扱いさせない（%USERPROFILE% を使うため）
    parser = configparser.ConfigParser(interpolation=None)
    if CONFIG_PATH.exists():
        parser.read(CONFIG_PATH, encoding="utf-8")
    g = parser["general"] if parser.has_section("general") else {}
    return {
        "watch_roots": _split_paths(g.get("watch_roots", str(TOOL_DIR.parent))),
        "folders": _split_paths(g.get("folders", str(TOOL_DIR.parent))),
        "mode": "ask" if str(g.get("mode", "auto")).strip().lower() == "ask" else "auto",
        "format": str(g.get("format", "html")).strip().lower(),
        "layout": "flat" if str(g.get("layout", "folder")).strip().lower() == "flat" else "folder",
        "flat_folders": {norm_dir(p) for p in _split_paths(g.get("flat_folders", ""))
                         if Path(p).is_dir()},
        "folder_folders": {norm_dir(p) for p in _split_paths(g.get("folder_folders", ""))
                           if Path(p).is_dir()},
        "original": str(g.get("original", "keep")).strip().lower(),
        "remote_images": str(g.get("remote_images", "true")).strip().lower() != "false",
        "poll_seconds": max(2, int(str(g.get("poll_seconds", "4")).strip() or 4)),
    }


def load_state():
    try:
        return json.loads(STATE_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def save_state(state):
    STATE_PATH.write_text(json.dumps(state, ensure_ascii=False, indent=1), encoding="utf-8")


def make_logger():
    def log(msg):
        line = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n"
        try:
            if LOG_PATH.exists() and LOG_PATH.stat().st_size > 1_000_000:
                LOG_PATH.replace(LOG_PATH.with_suffix(".log.old"))
            with open(LOG_PATH, "a", encoding="utf-8") as f:
                f.write(line)
        except OSError:
            pass
        if sys.stdout is not None:  # pythonw では画面出力なし
            try:
                print(line, end="")
            except OSError:
                pass
    return log


# ---------------------------------------------------------------- 変換処理の共通ルート

def is_excluded(path: Path, cfg) -> bool:
    """開発リポジトリ等、変換すべきでない場所かどうか。"""
    parts = set(path.parts)
    if parts & EXCLUDE_DIR_NAMES:
        return True
    if norm_dir(path.parent) == norm_dir(TOOL_DIR):
        return True
    # git リポジトリの作業ツリー内は対象外（上位フォルダーに .git があるか）
    roots = {norm_dir(r) for r in cfg["watch_roots"] if Path(r).is_dir()}
    d = path.parent
    for _ in range(30):
        if (d / ".git").exists():
            return True
        if norm_dir(d) in roots or d.parent == d:
            break
        d = d.parent
    return False


def try_convert(f: Path, cfg, state, log) -> str:
    """1 ファイルの変換を試みる。

    戻り値: converted(変換した) / seen(処理不要) / unconvertible(変換不可を記録) /
            declined(選択画面で見送り) / wait(書き込み中) / fail(エラーを記録)
    """
    try:
        key = str(f.resolve())
        st = f.stat()
    except OSError:
        return "seen"
    rec = state.get(key)
    if rec is not None:
        if rec.get("status") == "generated":
            return "seen"  # このツール自身が書き出したファイル
        if rec.get("size") == st.st_size and rec.get("mtime") == st.st_mtime:
            return "seen"  # 処理済みで変更なし
    if not file_ready(f):
        return "wait"
    try:
        r = convert_file(f, cfg, interactive=True)
        state[key] = {"size": st.st_size, "mtime": st.st_mtime,
                      "status": "ok", "output": str(r["out_dir"])}
        for gp in r["generated"]:
            state[str(gp.resolve())] = {"status": "generated"}
        where = r["out_dir"].name if r["layout"] == "folder" else f"{f.parent.name} に直接"
        log(f"変換完了: {f.name} → {where} "
            f"({'+'.join(r['saved'])}, 添付{len(r['attachments'])}件)")
        dispose_original(f, cfg, log)
        return "converted"
    except UserSkipped:
        state[key] = {"size": st.st_size, "mtime": st.st_mtime, "status": "declined"}
        log(f"選択画面で「変換しない」が選ばれました: {f.name}")
        return "declined"
    except Unconvertible as e:
        state[key] = {"size": st.st_size, "mtime": st.st_mtime,
                      "status": "skipped", "reason": str(e)}
        log(f"変換不可（元ファイルをそのまま残します）: {f.name} — {e}")
        return "unconvertible"
    except Exception:
        state[key] = {"size": st.st_size, "mtime": st.st_mtime, "status": "error"}
        log(f"変換エラー: {f.name}\n{traceback.format_exc()}")
        return "fail"


def file_ready(path: Path) -> bool:
    """書き込み途中のファイルに触らないための確認（サイズ安定＋排他オープン）。"""
    try:
        s1 = path.stat().st_size
    except OSError:
        return False
    time.sleep(1.0)
    try:
        if path.stat().st_size != s1:
            return False
        with open(path, "rb+"):
            pass
        return True
    except OSError:
        return False


def scan_once(cfg, state, log, folders=None) -> int:
    """指定フォルダー（既定: config の folders + 過去に変換実績のある場所）を巡回変換。"""
    converted = 0
    changed = False
    if folders is None:
        folders = list(cfg["folders"]) + known_dirs(state)
    seen = set()
    for folder in folders:
        folder = Path(folder)
        if norm_dir(folder) in seen or not folder.is_dir():
            continue
        seen.add(norm_dir(folder))
        for f in sorted(folder.glob("*.eml")) + sorted(folder.glob("*.msg")):
            result = try_convert(f, cfg, state, log)
            if result == "converted":
                converted += 1
            if result in ("converted", "unconvertible", "declined", "fail"):
                changed = True
    if changed:
        save_state(state)
    return converted


def known_dirs(state, limit=100):
    """過去に変換したファイルの置き場所（監視停止中のドロップの取りこぼし回収に使う）。"""
    dirs, seen = [], set()
    for key, rec in state.items():
        if rec.get("status") not in ("ok", "skipped", "error", "declined"):
            continue
        d = str(Path(key).parent)
        if d.lower() not in seen:
            seen.add(d.lower())
            dirs.append(d)
        if len(dirs) >= limit:
            break
    return dirs


# ---------------------------------------------------------------- 常駐監視

def acquire_lock():
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("127.0.0.1", LOCK_PORT))
        s.listen(1)
        return s
    except OSError:
        return None


def watch(cfg, log):
    lock = acquire_lock()
    if lock is None:
        log("既に別の監視プロセスが起動しているため終了します")
        return

    from watchdog.events import PatternMatchingEventHandler
    from watchdog.observers import Observer

    state = load_state()
    events = queue_mod.Queue()

    class Handler(PatternMatchingEventHandler):
        def __init__(self):
            super().__init__(patterns=["*.eml", "*.msg"], ignore_directories=True)

        def on_created(self, event):
            events.put(event.src_path)

        def on_moved(self, event):
            if str(event.dest_path).lower().endswith((".eml", ".msg")):
                events.put(event.dest_path)

    observer = Observer()
    watched = []
    for root in cfg["watch_roots"]:
        if Path(root).is_dir():
            observer.schedule(Handler(), root, recursive=True)
            watched.append(root)
        else:
            log(f"監視ルートが見つかりません: {root}")
    observer.daemon = True
    observer.start()
    log(f"監視を開始しました: {'; '.join(watched)} 配下のすべてのフォルダー "
        f"(モード: {cfg['mode']}, 保存スタイル: {cfg['layout']}, 形式: {cfg['format']})")

    # 起動直後に、これまでの置き場所の取りこぼしを回収
    try:
        n = scan_once(cfg, state, log)
        if n:
            log(f"起動時の取りこぼし回収: {n} 件変換")
    except Exception:
        log(f"起動時スキャンでエラー:\n{traceback.format_exc()}")

    pending = {}  # パス -> 試行回数（書き込み完了待ちのファイル）
    last_scan = time.monotonic()
    merged = False
    while True:
        try:
            # 新しいイベントを取り込む（最初の 1 件だけ最大 2 秒待ち、残りは即時回収）
            try:
                pending.setdefault(events.get(timeout=2.0), 0)
                while True:
                    pending.setdefault(events.get_nowait(), 0)
            except queue_mod.Empty:
                pass

            if pending and not merged:
                # 手動実行 (--once / --convert) が別プロセスで残した記録を合流させる
                disk = load_state()
                for k, v in disk.items():
                    state.setdefault(k, v)
                merged = True

            changed = False
            for p in list(pending):
                path = Path(p)
                if not path.exists() or path.suffix.lower() not in (".eml", ".msg"):
                    del pending[p]
                    continue
                if is_excluded(path, cfg):
                    del pending[p]
                    continue
                result = try_convert(path, cfg, state, log)
                if result == "wait":
                    pending[p] += 1
                    if pending[p] > 150:  # 約 5 分待っても書き込みが終わらない
                        log(f"書き込みが完了しないため保留を打ち切りました: {path.name}")
                        del pending[p]
                else:
                    del pending[p]
                    if result in ("converted", "unconvertible", "declined", "fail"):
                        changed = True
            if changed:
                save_state(state)
            if not pending:
                merged = False

            # 念のための定期巡回（イベントを取りこぼした場合の保険）
            if time.monotonic() - last_scan > max(30, cfg["poll_seconds"] * 5):
                scan_once(cfg, state, log)
                last_scan = time.monotonic()
        except Exception:
            log(f"監視ループでエラー:\n{traceback.format_exc()}")
            time.sleep(5)


# ---------------------------------------------------------------- main

def main(argv):
    cfg = load_config()
    log = make_logger()
    if not argv or argv[0] in ("-h", "--help"):
        print(__doc__)
        return 0
    if argv[0] == "--watch":
        watch(cfg, log)
        return 0
    if argv[0] == "--once":
        state = load_state()
        n = scan_once(cfg, state, log)
        log(f"一括変換が完了しました（{n} 件変換）")
        return 0
    if argv[0] == "--convert":
        if len(argv) < 2 or not Path(argv[1]).is_dir():
            log("--convert にはフォルダーを指定してください")
            return 1
        state = load_state()
        n = scan_once(cfg, state, log, folders=[argv[1]])
        log(f"{argv[1]} の変換が完了しました（{n} 件変換）")
        return 0
    # ファイル指定
    state = load_state()
    rc = 0
    for a in argv:
        p = Path(a)
        if not p.is_file():
            log(f"ファイルが見つかりません: {a}")
            rc = 1
            continue
        result = try_convert(p, cfg, state, log)
        if result == "fail":
            rc = 1
    save_state(state)
    return rc


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
