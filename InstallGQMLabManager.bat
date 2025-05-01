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

:: Create a temporary directory with a simple name to avoid path issues
set TEMP_DIR=C:\GQMTemp
echo Creating temporary directory at %TEMP_DIR%
if exist "%TEMP_DIR%" rmdir /s /q "%TEMP_DIR%"
mkdir "%TEMP_DIR%"
cd /d "%TEMP_DIR%"
echo Current directory: %CD%

:: Download the certificate using BITSAdmin (more reliable than PowerShell for some systems)
echo Downloading certificate...
bitsadmin /transfer "DownloadCertificate" https://gqmlab-solutions.github.io/msix-updates/GQMLab_Simple.cer "%TEMP_DIR%\GQMLab_Simple.cer"
if not exist "%TEMP_DIR%\GQMLab_Simple.cer" (
    echo Certificate download failed using bitsadmin, trying with certutil...
    certutil -urlcache -split -f https://gqmlab-solutions.github.io/msix-updates/GQMLab_Simple.cer "%TEMP_DIR%\GQMLab_Simple.cer"
    if not exist "%TEMP_DIR%\GQMLab_Simple.cer" (
        echo All certificate download attempts failed.
        echo Please try manual installation instead.
        pause
        exit /b
    )
)

:: Download the MSIX package
echo Downloading MSIX package...
bitsadmin /transfer "DownloadMSIX" https://gqmlab-solutions.github.io/msix-updates/gqm_lab_manager.msix "%TEMP_DIR%\gqm_lab_manager.msix"
if not exist "%TEMP_DIR%\gqm_lab_manager.msix" (
    echo MSIX download failed using bitsadmin, trying with certutil...
    certutil -urlcache -split -f https://gqmlab-solutions.github.io/msix-updates/gqm_lab_manager.msix "%TEMP_DIR%\gqm_lab_manager.msix"
    if not exist "%TEMP_DIR%\gqm_lab_manager.msix" (
        echo All MSIX download attempts failed.
        echo Please try manual installation instead.
        pause
        exit /b
    )
)

:: Install the certificate to Root store
echo Installing certificate...
certutil -addstore -f Root "%TEMP_DIR%\GQMLab_Simple.cer"
if %errorLevel% neq 0 (
    echo Certificate installation failed.
    echo Please try manual installation instead.
    pause
    exit /b
)

echo Certificate installed successfully.
echo.

:: Install the MSIX package
echo Installing GQM Lab Manager...
powershell -Command "Add-AppxPackage -Path '%TEMP_DIR%\gqm_lab_manager.msix'"
if %errorLevel% neq 0 (
    echo Installation failed.
    echo Please try manual installation instead.
    pause
    exit /b
)

echo.
echo GQM Lab Manager installed successfully!
echo.
echo You can now launch the application by:
echo  1. Clicking the Start button
echo  2. Typing "GQM Lab Manager"
echo  3. Clicking on the GQM Lab Manager app icon
echo.
echo Press any key to exit...
pause

:: Clean up
rmdir /s /q "%TEMP_DIR%"