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

:: Create a temporary directory
set TEMP_DIR=%TEMP%\GQMLabInstall
mkdir "%TEMP_DIR%" 2>nul

echo Installing certificate...
:: Check if certificate exists locally first
if exist "GQMLab_Simple.cer" (
    echo Using local certificate...
    copy "GQMLab_Simple.cer" "%TEMP_DIR%\GQMLab_Simple.cer" >nul
) else (
    :: Try to download it
    echo Local certificate not found, downloading...
    powershell -Command "(New-Object System.Net.WebClient).DownloadFile('https://gqmlab-solutions.github.io/msix-updates/GQMLab_Simple.cer', '%TEMP_DIR%\GQMLab_Simple.cer')"
    if %errorLevel% neq 0 (
        echo Failed to download certificate.
        pause
        exit /b
    )
)

:: Install certificate to BOTH Trusted Root and Trusted People stores
echo Installing certificate to Trusted Root store...
certutil -addstore Root "%TEMP_DIR%\GQMLab_Simple.cer"
if %errorLevel% neq 0 (
    echo Certificate installation to Root store failed.
    pause
    exit /b
)

echo Installing certificate to Trusted People store...
certutil -addstore TrustedPeople "%TEMP_DIR%\GQMLab_Simple.cer"
if %errorLevel% neq 0 (
    echo Certificate installation to Trusted People store failed.
    pause
    exit /b
)

echo Certificate installed successfully.
echo.

echo Installing GQM Lab Manager...
:: Check if MSIX exists locally
if exist "gqm_lab_manager.msix" (
    echo Using local MSIX package...
    copy "gqm_lab_manager.msix" "%TEMP_DIR%\gqm_lab_manager.msix" >nul
) else (
    :: Try to download it
    echo Local MSIX not found, downloading...
    powershell -Command "(New-Object System.Net.WebClient).DownloadFile('https://gqmlab-solutions.github.io/msix-updates/gqm_lab_manager.msix', '%TEMP_DIR%\gqm_lab_manager.msix')"
    if %errorLevel% neq 0 (
        echo Failed to download MSIX package.
        pause
        exit /b
    )
)

:: Add developer mode option to bypass some certificate issues
echo Enabling developer mode features...
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock" /t REG_DWORD /f /v "AllowDevelopmentWithoutDevLicense" /d "1"

:: Install with force option
echo Installing package with force option...
powershell -Command "Add-AppxPackage -Path '%TEMP_DIR%\gqm_lab_manager.msix' -ForceApplicationShutdown -ForceUpdateFromAnyVersion -DevelopmentMode"
if %errorLevel% neq 0 (
    echo First installation attempt failed, trying alternative method...
    
    :: Try direct install as fallback
    powershell -Command "Add-AppxPackage -Path '%TEMP_DIR%\gqm_lab_manager.msix' -ForceApplicationShutdown -ForceUpdateFromAnyVersion"
    
    if %errorLevel% neq 0 (
        echo All installation attempts failed.
        echo Please contact support for assistance.
        echo The MSIX file is available at: %TEMP_DIR%\gqm_lab_manager.msix
        pause
        exit /b
    }
)

echo.
echo GQM Lab Manager installed successfully!
echo.
echo You can now launch the application from the Start Menu.
echo.
pause

:: Clean up
rmdir /s /q "%TEMP_DIR%"