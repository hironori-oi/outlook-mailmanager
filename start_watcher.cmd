@echo off
chcp 932 >nul
rem メール変換の常駐監視を起動する（既に動いていれば二重起動しない）
setlocal
set "PYW="
for /f "delims=" %%i in ('where pythonw 2^>nul') do if not defined PYW set "PYW=%%i"
if not defined PYW (
    echo pythonw.exe が見つかりません。Python をインストールしてください。
    pause
    exit /b 1
)
start "" "%PYW%" "%~dp0mail_converter.py" --watch
echo 監視プログラムを起動しました（既に動いていた場合はそのまま）
timeout /t 3 >nul
