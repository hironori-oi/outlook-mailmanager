@echo off
chcp 932 >nul
rem メール変換の常駐監視を停止する（自動起動の登録は残す。再開は start_watcher.cmd）
taskkill /f /im OutlookMailConverter.exe >nul 2>&1
powershell -NoProfile -Command "Get-CimInstance Win32_Process | Where-Object { $_.Name -match '^python' -and $_.CommandLine -match 'mail_converter' -and $_.ProcessId -ne $PID } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }"
echo 監視プログラムを停止しました
timeout /t 3 >nul
