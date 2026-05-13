@echo off
chcp 65001 >nul 2>&1
echo.
echo ================================================
echo   KASA DASHBOARD - Tum Windows Sunuculari Durdur
echo ================================================
echo.
echo Sinyal gonderiliyor...
echo.  > "\\YOUR-FILE-SERVER\SharedFolders\RAPOR\Kasa\dist\stop.flag"
echo [OK] Sinyal gonderildi.
echo [OK] Tum Windows PC'lerdeki sunucular ~10 saniye icerisinde kapanacak.
echo [OK] Ubuntu sunucusu (YOUR-UBUNTU-SERVER-IP:8765) calismaya devam edecek.
echo.
pause
