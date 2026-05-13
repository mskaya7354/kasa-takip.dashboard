@echo off
chcp 65001 >nul 2>&1
echo.
echo ================================================
echo   KASA DASHBOARD - Yerel Sunucuyu Durdur
echo ================================================
echo.

taskkill /F /IM pythonw.exe >nul 2>&1
taskkill /F /IM python.exe /FI "WINDOWTITLE eq Kasa*" >nul 2>&1

for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":8765 " ^| findstr "LISTENING" 2^>nul') do (
    taskkill /PID %%a /F >nul 2>&1
)

echo [OK] Yerel sunucu durduruldu.
echo [OK] Ubuntu sunucusu (YOUR-UBUNTU-SERVER-IP:8765) calismaya devam ediyor.
echo.
pause
