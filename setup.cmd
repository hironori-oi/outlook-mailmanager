@echo off
chcp 932 >nul
rem ============================================================
rem メール変換ツールのセットアップ（新しい PC への移植用）
rem このファイルをダブルクリックすると:
rem   1. 必要な Python ライブラリを導入
rem   2. サインイン時の自動起動を登録
rem   3. 監視プログラムを今すぐ起動
rem ============================================================
setlocal

rem --- Python の確認
set "PY="
for /f "delims=" %%i in ('where python 2^>nul') do if not defined PY set "PY=%%i"
if not defined PY (
    echo [エラー] Python が見つかりません。
    echo Microsoft Store で「Python 3.13」を検索してインストールしてから、
    echo もう一度このファイルをダブルクリックしてください。
    pause
    exit /b 1
)
echo Python: %PY%

set "PYW="
for /f "delims=" %%i in ('where pythonw 2^>nul') do if not defined PYW set "PYW=%%i"
if not defined PYW (
    echo [エラー] pythonw.exe が見つかりません。Python を入れ直してください。
    pause
    exit /b 1
)

rem --- ライブラリの導入
echo 必要なライブラリを導入しています...
"%PY%" -m pip install --quiet extract_msg watchdog customtkinter pillow
if errorlevel 1 (
    echo [エラー] ライブラリの導入に失敗しました。インターネット接続を確認してください。
    pause
    exit /b 1
)

rem --- サインイン時の自動起動を登録（スタートアップにショートカット作成）
powershell -NoProfile -Command ^
  "$ws = New-Object -ComObject WScript.Shell;" ^
  "$lnk = $ws.CreateShortcut(\"$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\OutlookMailConverter.lnk\");" ^
  "$lnk.TargetPath = '%PYW%';" ^
  "$lnk.Arguments = '\"%~dp0mail_converter.py\" --watch';" ^
  "$lnk.WorkingDirectory = '%~dp0';" ^
  "$lnk.Description = 'Outlook mail auto converter';" ^
  "$lnk.Save()"
echo 自動起動を登録しました。

rem --- 今すぐ起動
start "" "%PYW%" "%~dp0mail_converter.py" --watch
echo.
echo セットアップ完了。監視プログラムが起動しました。
echo 動作の記録は converter.log、設定変更は config.ini をご覧ください。
pause
