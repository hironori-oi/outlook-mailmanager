@echo off
chcp 932 >nul
rem ƒ[ƒ‹•ÏŠ·‚Ìí’“ŠÄŽ‹‚ð’âŽ~‚·‚é
powershell -NoProfile -Command "Get-CimInstance Win32_Process | Where-Object { $_.Name -match '^python' -and $_.CommandLine -match 'mail_converter' -and $_.ProcessId -ne $PID } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }"
echo ŠÄŽ‹ƒvƒƒOƒ‰ƒ€‚ð’âŽ~‚µ‚Ü‚µ‚½
timeout /t 3 >nul
