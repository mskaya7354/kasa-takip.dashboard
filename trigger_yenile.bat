@echo off
chcp 65001 >nul 2>&1
echo.
echo ================================================
echo   KASA DASHBOARD - Tum Bilgisayarlari Yenile
echo ================================================
echo.
echo Sinyal gonderiliyor...
echo.  > "\\YOUR-FILE-SERVER\SharedFolders\RAPOR\Kasa\dist\restart.flag"
echo [OK] Sinyal gonderildi.
echo.
echo Tum acik bilgisayarlardaki dashboard ~10 saniye
echo icerisinde otomatik olarak yeniden baslar.
echo.
pause
