@echo off
echo GQM Lab Manager Installer
echo ========================
echo.

:: Create a log file
set LOG_FILE=%USERPROFILE%\Desktop\GQMLabInstallLog.txt
echo GQM Lab Manager Installation Log > %LOG_FILE%
echo Started at: %DATE% %TIME% >> %LOG_FILE%
echo. >> %LOG_FILE%

:: Check if running as administrator
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo This installer needs to be run as administrator.
    echo Please right-click this file and select "Run as administrator".
    echo. 
    echo Installation aborted. Please run as administrator. >> %LOG_FILE%
    pause
    exit /b
)

echo Running as administrator: YES >> %LOG_FILE%

:: Create a temporary directory
set TEMP_DIR=%TEMP%\GQMLabInstall
mkdir "%TEMP_DIR%" 2>nul

echo Preparing for installation...
echo Current directory: %CD% >> %LOG_FILE%

:: Step 1: Download files if needed (just like in manual installation)
echo Checking for certificate and MSIX files...

:: First check if files exist locally
if exist "GQMLab_Simple.cer" (
    echo Found certificate locally. >> %LOG_FILE%
    copy "GQMLab_Simple.cer" "%TEMP_DIR%\GQMLab_Simple.cer" >nul
) else (
    echo Downloading certificate... >> %LOG_FILE%
    powershell -Command "Invoke-WebRequest -Uri 'https://gqmlab-solutions.github.io/msix-updates/GQMLab_Simple.cer' -OutFile '%TEMP_DIR%\GQMLab_Simple.cer'"
    if %errorLevel% neq 0 (
        echo Failed to download certificate. >> %LOG_FILE%
        echo Failed to download certificate.
        echo Please try manual installation instead.
        pause
        exit /b
    )
)

if exist "gqm_lab_manager.msix" (
    echo Found MSIX package locally. >> %LOG_FILE%
    copy "gqm_lab_manager.msix" "%TEMP_DIR%\gqm_lab_manager.msix" >nul
) else (
    echo Downloading MSIX package... >> %LOG_FILE%
    powershell -Command "Invoke-WebRequest -Uri 'https://gqmlab-solutions.github.io/msix-updates/gqm_lab_manager.msix' -OutFile '%TEMP_DIR%\gqm_lab_manager.msix'"
    if %errorLevel% neq 0 (
        echo Failed to download MSIX package. >> %LOG_FILE%
        echo Failed to download MSIX package.
        echo Please try manual installation instead.
        pause
        exit /b
    )
)

:: Step 2: Install the certificate exactly as in manual process
echo Installing certificate...
echo Installing certificate to Trusted Root store (same as manual installation)... >> %LOG_FILE%

:: This mimics the manual process of installing to Trusted Root
certutil -addstore -f Root "%TEMP_DIR%\GQMLab_Simple.cer" >> %LOG_FILE% 2>&1
if %errorLevel% neq 0 (
    echo Certificate installation failed. >> %LOG_FILE%
    echo Certificate installation failed.
    echo Please try manual installation instead.
    pause
    exit /b
)

echo Certificate installed successfully. >> %LOG_FILE%
echo Certificate installed successfully.
echo.

:: Step 3: Install the MSIX package directly (as in manual process)
echo Installing GQM Lab Manager...
echo Installing MSIX package (same as double-clicking it)... >> %LOG_FILE%

:: This is similar to double-clicking the MSIX file
powershell -Command "Add-AppxPackage -Path '%TEMP_DIR%\gqm_lab_manager.msix'" >> %LOG_FILE% 2>&1
if %errorLevel% neq 0 (
    echo Installation failed. >> %LOG_FILE%
    echo Installation failed.
    echo Please try manual installation instead.
    echo Details have been saved to: %LOG_FILE%
    pause
    exit /b
)

:: Check if app was installed successfully
echo Verifying installation... >> %LOG_FILE%
powershell -Command "$app = Get-AppxPackage | Where-Object {$_.Name -like '*gqm*' -or $_.Name -like '*lab*'}; if ($app) { Write-Host 'App installed: ' + $app.Name; exit 0 } else { Write-Host 'App not found after installation'; exit 1 }" >> %LOG_FILE% 2>&1

if %errorLevel% neq 0 (
    echo Warning: Could not verify app installation. >> %LOG_FILE%
) else (
    echo App installation verified. >> %LOG_FILE%
    
    :: Create a desktop shortcut for convenience
    echo Creating desktop shortcut... >> %LOG_FILE%
    powershell -Command "$app = Get-AppxPackage | Where-Object {$_.Name -like '*gqm*' -or $_.Name -like '*lab*'}; if ($app) { $appId = $app.PackageFamilyName; $Shell = New-Object -ComObject WScript.Shell; $Shortcut = $Shell.CreateShortcut('%USERPROFILE%\Desktop\GQM Lab Manager.lnk'); $Shortcut.TargetPath = 'shell:AppsFolder\' + $appId; $Shortcut.Save(); Write-Host 'Desktop shortcut created.' }" >> %LOG_FILE% 2>&1
)

echo.
echo GQM Lab Manager installed successfully!
echo.
echo You can now launch the application by:
echo  1. Clicking the Start button
echo  2. Typing "GQM Lab Manager"
echo  3. Clicking on the GQM Lab Manager app icon
echo.
echo A shortcut may also have been created on your Desktop.
echo.
echo If you can't find the app, please try the manual installation method
echo from our website.
echo.
echo Press any key to exit...
pause > nul

:: Clean up
echo Cleaning up temporary files >> %LOG_FILE%
rmdir /s /q "%TEMP_DIR%" >> %LOG_FILE% 2>&1
echo Installation process completed at: %DATE% %TIME% >> %LOG_FILE%