@echo off
chcp 65001 >nul 2>&1

:: Yonetici degil ise UAC ile yeniden baslat
net session >nul 2>&1
if %errorLevel% neq 0 (
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

echo.
echo ================================================
echo   KASA DASHBOARD - PC Kurulumu
echo ================================================
echo.

set "STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "VBS_PATH=%~dp0gizli.vbs"
set "LOCAL_PY=C:\KasaDashboard\python"
set "NET_PY=%~dp0..\_ARSIV\python"

echo [1/3] Python yerel diske kopyalaniyor (210 MB, bir kez yapilir)...
if exist "%LOCAL_PY%\pythonw.exe" (
    echo [OK] Zaten kurulu, atlanıyor.
) else (
    robocopy "%NET_PY%" "%LOCAL_PY%" /E /NP /NDL
    echo [OK] Python kopyalandi: %LOCAL_PY%
)

echo [1b/3] zeroconf paketi kuruluyor (mDNS / kasa.local destegi)...
"%LOCAL_PY%\python.exe" -m pip install zeroconf --quiet 2>nul
echo [OK] Paket hazir.

echo [2/3] Baslangica ekleniyor...
powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%STARTUP%\KasaDashboard.lnk'); $s.TargetPath = 'wscript.exe'; $s.Arguments = '\"%VBS_PATH%\"'; $s.WorkingDirectory = '%~dp0'; $s.WindowStyle = 7; $s.Save()"
echo [OK] Bilgisayar acilisinda otomatik baslar.

echo [3/3] Guvenlik duvari kurali aciliyor (port 8765)...
powershell -Command "if (-not (Get-NetFirewallRule -DisplayName 'Kasa Dashboard 8765' -ErrorAction SilentlyContinue)) { New-NetFirewallRule -DisplayName 'Kasa Dashboard 8765' -Direction Inbound -Protocol TCP -LocalPort 8765 -Action Allow | Out-Null; Write-Host '[OK] Kural eklendi.' } else { Write-Host '[OK] Kural zaten mevcut.' }"

echo [4/4] Simdi baslatiliyor...
wscript "%VBS_PATH%"
echo [OK] Sunucu arka planda baslatildi.
echo.
echo Tarayici sunucu hazir oldugunda otomatik acilacak.
echo.
pause
