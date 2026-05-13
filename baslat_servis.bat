@echo off
chcp 65001 >nul 2>&1

rem Startup kisayolunu kalici olarak kaldir
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\KasaDashboard.lnk" >nul 2>&1

rem Calisiyorsa durdur
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":8765 " ^| findstr "LISTENING" 2^>nul') do (
    taskkill /PID %%a /F >nul 2>&1
)
taskkill /F /IM pythonw.exe >nul 2>&1

rem Sessizce cik - bir daha calisma
exit /b 0
