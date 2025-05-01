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
certutil -addstore Root GQMLab_Simple.cer
if %errorLevel% neq 0 (
    echo Certificate installation failed.
    pause
    exit /b
)
echo Certificate installed successfully.
echo.

echo Installing GQM Lab Manager...
powershell -Command "Add-AppxPackage -Path 'gqm_lab_manager.msix'"
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