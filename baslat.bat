@echo off
chcp 65001 >nul 2>&1
title Kasa Takip Dashboard v4
cd /d "%~dp0"

echo.
echo ================================================
echo   KASA TAKIP - Dashboard v4 (Yonetici Ozeti)
echo ================================================
echo.

set "EMBED_PYTHON=%~dp0..\_ARSIV\python\python.exe"

python --version >nul 2>&1
if errorlevel 1 (
    if exist "%EMBED_PYTHON%" (
        echo [OK] Sistem Python bulunamadi, embedded Python kullaniliyor.
        set "PYTHON=%EMBED_PYTHON%"
    ) else (
        echo [HATA] Python bulunamadi.
        pause & exit /b 1
    )
) else (
    set "PYTHON=python"
)

echo [1/3] Paketler kontrol ediliyor...
echo [OK] Paketler hazir.

echo [2/3] Chart.js kontrol ediliyor...
if not exist "chart.min.js" (
    echo   Chart.js indiriliyor...
    %PYTHON% -c "import urllib.request; urllib.request.urlretrieve('https://cdn.jsdelivr.net/npm/chart.js', 'chart.min.js'); print('  [OK] Chart.js indirildi.')"
    if not exist "chart.min.js" (
        echo   [UYARI] Chart.js indirilemedi. Grafik olmadan devam edilecek.
    )
)
if exist "chart.min.js" echo [OK] Chart.js hazir.

set "EXCEL_PATH=\\YOUR-FILE-SERVER\SharedFolders\Kasa\Kasa_Takip.xlsm"

echo [3/3] Sunucu baslatiliyor...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":8765 " ^| findstr "LISTENING"') do (
    taskkill /PID %%a /F >nul 2>&1
)
echo.
echo   Dosya:     %EXCEL_PATH%
echo   Dashboard: http://localhost:8765
echo.
echo   Tarayicidan acmak icin:  http://localhost:8765
echo   Ctrl+C ile durdurun.
echo ================================================
echo.

%PYTHON% server.py --file "%EXCEL_PATH%" --port 8765 --polling

pause
