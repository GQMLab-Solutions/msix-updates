@echo off
echo GQM Lab Manager Installer
echo ========================
echo.

:: Check if running as administrator
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo This installer needs to be run as administrator.
    echo Please right-click this file and select "Run as administrator".
    echo.
    pause
    exit /b
)

echo Installing certificate...
:: Check if certificate exists locally first
if exist "GQMLab_Simple.cer" (
    certutil -addstore Root "GQMLab_Simple.cer"
) else (
    :: Try to download it
    echo Local certificate not found, downloading...
    powershell -Command "(New-Object System.Net.WebClient).DownloadFile('https://gqmlab-solutions.github.io/msix-updates/GQMLab_Simple.cer', '%TEMP%\GQMLab_Simple.cer')"
    certutil -addstore Root "%TEMP%\GQMLab_Simple.cer"
)

if %errorLevel% neq 0 (
    echo Certificate installation failed.
    pause
    exit /b
)
echo Certificate installed successfully.
echo.

echo Installing GQM Lab Manager...
:: Check if MSIX exists locally
if exist "gqm_lab_manager.msix" (
    powershell -Command "Add-AppxPackage -Path 'gqm_lab_manager.msix'"
) else (
    :: Try to download it
    echo Local MSIX not found, downloading...
    powershell -Command "(New-Object System.Net.WebClient).DownloadFile('https://gqmlab-solutions.github.io/msix-updates/gqm_lab_manager.msix', '%TEMP%\gqm_lab_manager.msix')"
    powershell -Command "Add-AppxPackage -Path '%TEMP%\gqm_lab_manager.msix'"
)

if %errorLevel% neq 0 (
    echo Installation failed.
    pause
    exit /b
)

echo.
echo GQM Lab Manager installed successfully!
echo.
echo You can now launch the application from the Start Menu.
echo.
pause