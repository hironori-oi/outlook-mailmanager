@echo off
chcp 932 >nul
rem ============================================================
rem メール変換ツールを完全に停止する（自動起動も解除）
rem また使いたくなったら OutlookMailConverter.exe をダブルクリック
rem ============================================================
taskkill /f /im OutlookMailConverter.exe >nul 2>&1
powershell -NoProfile -Command "Get-CimInstance Win32_Process | Where-Object { $_.Name -match '^python' -and $_.CommandLine -match 'mail_converter' -and $_.ProcessId -ne $PID } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }"
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\OutlookMailConverter.lnk" >nul 2>&1
echo メール変換ツールを停止し、自動起動も解除しました。
echo また使うときは OutlookMailConverter.exe をダブルクリックしてください。
pause
