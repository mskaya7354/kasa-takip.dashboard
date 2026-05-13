@echo off
cd /d "%~dp0"
echo.
echo Kasa Takip Dashboard
echo.
set "EXCEL_PATH=Z:\RAPOR\Kasa\Kasa_Takip_v3.xlsm"
echo   Dosya: %EXCEL_PATH%
echo   Dashboard: http://localhost:8765
echo.
python server.py --file "%EXCEL_PATH%" --port 8765 --polling
pause
