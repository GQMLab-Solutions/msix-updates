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

echo Preparing for installation...
:: Debug info
echo Current directory: %CD%
dir /b

echo Installing certificate...
:: First try with local cer file (check all possible names to handle case sensitivity)
set CERT_FOUND=0
if exist "GQMLab_Simple.cer" (
    echo Found: GQMLab_Simple.cer
    copy "GQMLab_Simple.cer" "%TEMP_DIR%\cert.cer" >nul
    set CERT_FOUND=1
) else if exist "gqmlab_simple.cer" (
    echo Found: gqmlab_simple.cer
    copy "gqmlab_simple.cer" "%TEMP_DIR%\cert.cer" >nul
    set CERT_FOUND=1
) else if exist "GQMlab_simple.cer" (
    echo Found: GQMlab_simple.cer
    copy "GQMlab_simple.cer" "%TEMP_DIR%\cert.cer" >nul
    set CERT_FOUND=1
)

:: If not found locally, download it
if %CERT_FOUND%==0 (
    echo Certificate not found locally, downloading...
    powershell -Command "Invoke-WebRequest -Uri 'https://gqmlab-solutions.github.io/msix-updates/GQMLab_Simple.cer' -OutFile '%TEMP_DIR%\cert.cer'"
    if %errorLevel% neq 0 (
        echo Failed to download certificate.
        echo Trying alternative URL...
        powershell -Command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/gqmlab-solutions/msix-updates/main/GQMLab_Simple.cer' -OutFile '%TEMP_DIR%\cert.cer'"
        if %errorLevel% neq 0 (
            echo All download attempts failed.
            pause
            exit /b
        )
    )
)

:: Install certificate to BOTH Trusted Root and Trusted People stores
echo Installing certificate to Trusted Root store...
certutil -addstore Root "%TEMP_DIR%\cert.cer"
if %errorLevel% neq 0 (
    echo Certificate installation to Root store failed.
    pause
    exit /b
)

echo Installing certificate to Trusted People store...
certutil -addstore TrustedPeople "%TEMP_DIR%\cert.cer"
if %errorLevel% neq 0 (
    echo Certificate installation to Trusted People store failed.
    pause
    exit /b
)

echo Certificate installed successfully.
echo.

echo Installing GQM Lab Manager...
:: Check if MSIX exists locally
set MSIX_FOUND=0
if exist "gqm_lab_manager.msix" (
    echo Found: gqm_lab_manager.msix
    copy "gqm_lab_manager.msix" "%TEMP_DIR%\app.msix" >nul
    set MSIX_FOUND=1
) else if exist "GQM_Lab_Manager.msix" (
    echo Found: GQM_Lab_Manager.msix
    copy "GQM_Lab_Manager.msix" "%TEMP_DIR%\app.msix" >nul
    set MSIX_FOUND=1
)

:: If not found locally, download it
if %MSIX_FOUND%==0 (
    echo MSIX package not found locally, downloading...
    powershell -Command "Invoke-WebRequest -Uri 'https://gqmlab-solutions.github.io/msix-updates/gqm_lab_manager.msix' -OutFile '%TEMP_DIR%\app.msix'"
    if %errorLevel% neq 0 (
        echo Failed to download MSIX package.
        echo Trying alternative URL...
        powershell -Command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/gqmlab-solutions/msix-updates/main/gqm_lab_manager.msix' -OutFile '%TEMP_DIR%\app.msix'"
        if %errorLevel% neq 0 (
            echo All download attempts failed.
            pause
            exit /b
        )
    )
)

:: Add developer mode option to bypass some certificate issues
echo Enabling developer mode features...
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock" /t REG_DWORD /f /v "AllowDevelopmentWithoutDevLicense" /d "1"

:: Try multiple installation methods
echo Installing package (method 1)...
powershell -Command "Add-AppxPackage -Path '%TEMP_DIR%\app.msix' -ForceApplicationShutdown -ForceUpdateFromAnyVersion"
if %errorLevel% neq 0 (
    echo Method 1 failed, trying method 2...
    powershell -Command "Add-AppxPackage -Path '%TEMP_DIR%\app.msix' -ForceApplicationShutdown -ForceUpdateFromAnyVersion -DevelopmentMode"
    
    if %errorLevel% neq 0 (
        echo Method 2 failed, trying method 3...
        :: Try with direct certutil import
        echo Importing certificate directly from the MSIX package...
        mkdir "%TEMP_DIR%\extract" 2>nul
        powershell -Command "Expand-Archive -Path '%TEMP_DIR%\app.msix' -DestinationPath '%TEMP_DIR%\extract' -Force"
        
        :: Find and install any certificates in the package
        for /r "%TEMP_DIR%\extract" %%f in (*.cer) do (
            echo Found certificate: %%f
            certutil -addstore Root "%%f"
            certutil -addstore TrustedPeople "%%f"
        )
        
        :: Try install again
        powershell -Command "Add-AppxPackage -Path '%TEMP_DIR%\app.msix' -ForceUpdateFromAnyVersion"
        
        if %errorLevel% neq 0 (
            echo All installation attempts failed.
            echo.
            echo Please try the following:
            echo 1. Make sure Windows is up to date
            echo 2. Contact support for assistance
            echo.
            echo Technical details:
            echo The MSIX file is available at: %TEMP_DIR%\app.msix
            echo The certificate is available at: %TEMP_DIR%\cert.cer
            pause
            exit /b
        }
    }
)

echo.
echo GQM Lab Manager installed successfully!
echo.
echo You can now launch the application by:
echo  1. Clicking the Start button
echo  2. Typing "GQM Lab Manager"
echo  3. Clicking on the GQM Lab Manager app icon
echo.
echo You can also pin the app to your taskbar or Start menu for easier access.
echo.
pause

:: Clean up
rmdir /s /q "%TEMP_DIR%"