@echo off
color 1F
title GQM Lab Manager - Installation Wizard
cls

echo ================================================================================================
echo                            GQM LAB MANAGER - INSTALLATION WIZARD
echo ================================================================================================
echo.
echo  Welcome to the GQM Lab Manager installation wizard. This installer will guide you through
echo  the process of setting up GQM Lab Manager on your computer.
echo.
echo  This wizard will:
echo    1. Install the required security certificate
echo    2. Install the GQM Lab Manager application
echo    3. Create shortcuts for easy access
echo.
echo  Please ensure you're connected to the internet during this process.
echo.
echo ================================================================================================
echo.
pause
cls

:: Check if running as administrator
echo ================================================================================================
echo                                   CHECKING PERMISSIONS
echo ================================================================================================
echo.
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo  ERROR: This installer needs to be run as an administrator.
    echo.
    echo  Please right-click on InstallGQMLabManager.bat and select "Run as administrator",
    echo  then try again.
    echo.
    echo ================================================================================================
    echo.
    pause
    exit /b
)
echo  [✓] Running with administrator permissions.
echo.

:: Create a temporary directory with a simple name to avoid path issues
set TEMP_DIR=C:\GQMTemp
echo  Creating temporary directory...
if exist "%TEMP_DIR%" rmdir /s /q "%TEMP_DIR%" >nul
mkdir "%TEMP_DIR%" >nul
cd /d "%TEMP_DIR%"
echo  [✓] Temporary directory created.
echo.

echo ================================================================================================
echo                                 DOWNLOADING REQUIRED FILES
echo ================================================================================================
echo.
echo  Step 1 of 3: Downloading certificate...
echo.
bitsadmin /transfer /download "DownloadCertificate" https://gqmlab-solutions.github.io/msix-updates/GQMLab_Simple.cer "%TEMP_DIR%\GQMLab_Simple.cer" >nul
if not exist "%TEMP_DIR%\GQMLab_Simple.cer" (
    echo  [!] Certificate download failed, trying alternative method...
    certutil -urlcache -split -f https://gqmlab-solutions.github.io/msix-updates/GQMLab_Simple.cer "%TEMP_DIR%\GQMLab_Simple.cer" >nul
    if not exist "%TEMP_DIR%\GQMLab_Simple.cer" (
        echo  ERROR: Could not download certificate.
        echo.
        echo  This could be due to:
        echo   - Internet connection issues
        echo   - Firewall or security software blocking the download
        echo.
        echo  Please try the All-in-One Package from our website instead:
        echo  https://gqmlab-solutions.github.io/msix-updates/
        echo.
        echo ================================================================================================
        echo.
        pause
        exit /b
    )
)
echo  [✓] Certificate downloaded successfully.
echo.

echo  Step 2 of 3: Downloading application package...
echo.
bitsadmin /transfer /download "DownloadMSIX" https://gqmlab-solutions.github.io/msix-updates/gqm_lab_manager.msix "%TEMP_DIR%\gqm_lab_manager.msix" >nul
if not exist "%TEMP_DIR%\gqm_lab_manager.msix" (
    echo  [!] Application download failed, trying alternative method...
    certutil -urlcache -split -f https://gqmlab-solutions.github.io/msix-updates/gqm_lab_manager.msix "%TEMP_DIR%\gqm_lab_manager.msix" >nul
    if not exist "%TEMP_DIR%\gqm_lab_manager.msix" (
        echo  ERROR: Could not download application package.
        echo.
        echo  This could be due to:
        echo   - Internet connection issues
        echo   - Firewall or security software blocking the download
        echo.
        echo  Please try the All-in-One Package from our website instead:
        echo  https://gqmlab-solutions.github.io/msix-updates/
        echo.
        echo ================================================================================================
        echo.
        pause
        exit /b
    )
)
echo  [✓] Application package downloaded successfully.
echo.
echo ================================================================================================
echo                                INSTALLING CERTIFICATE
echo ================================================================================================
echo.
echo  Installing security certificate...
echo  This certificate allows Windows to trust software from Geoquip Marine.
echo.
certutil -addstore -f Root "%TEMP_DIR%\GQMLab_Simple.cer" >nul
if %errorLevel% neq 0 (
    echo  ERROR: Certificate installation failed.
    echo.
    echo  This could be due to:
    echo   - Security software blocking the installation
    echo   - Windows policy restrictions
    echo.
    echo  Please try the manual installation method from our website:
    echo  https://gqmlab-solutions.github.io/msix-updates/
    echo.
    echo ================================================================================================
    echo.
    pause
    exit /b
)
echo  [✓] Certificate installed successfully.
echo.

echo ================================================================================================
echo                              INSTALLING GQM LAB MANAGER
echo ================================================================================================
echo.
echo  Step 3 of 3: Installing GQM Lab Manager...
echo  The application is being installed on your computer. This may take a moment.
echo.
powershell -Command "Add-AppxPackage -Path '%TEMP_DIR%\gqm_lab_manager.msix'" >nul
if %errorLevel% neq 0 (
    echo  ERROR: Installation failed.
    echo.
    echo  This could be due to:
    echo   - Security software blocking the installation
    echo   - Incompatible system configuration
    echo   - Certificate issues
    echo.
    echo  Please try the manual installation method from our website:
    echo  https://gqmlab-solutions.github.io/msix-updates/
    echo.
    echo ================================================================================================
    echo.
    pause
    exit /b
)

echo  [✓] GQM Lab Manager installed successfully!
echo.

:: Create desktop shortcut
echo  Creating desktop shortcut...
powershell -Command "$packages = Get-AppxPackage | Where-Object {$_.Name -like '*gqm*' -or $_.PackageFamilyName -like '*gqm*' -or $_.Name -like '*lab*'}; if ($packages) { $appId = $packages[0].PackageFamilyName; $Shell = New-Object -ComObject WScript.Shell; $Shortcut = $Shell.CreateShortcut([System.IO.Path]::Combine($env:USERPROFILE, 'Desktop', 'GQM Lab Manager.lnk')); $Shortcut.TargetPath = 'shell:AppsFolder\' + $appId; $Shortcut.Save(); Write-Host '  [✓] Desktop shortcut created.' }" 2>nul
echo.

:: Clean up
echo  Cleaning up temporary files...
cd /d "%USERPROFILE%"
rmdir /s /q "%TEMP_DIR%" >nul 2>&1
echo  [✓] Cleanup complete.
echo.

echo ================================================================================================
echo                              INSTALLATION COMPLETE!
echo ================================================================================================
echo.
echo  GQM Lab Manager has been successfully installed on your computer!
echo.
echo  You can start the application using any of these methods:
echo.
echo    1. Double-click the GQM Lab Manager shortcut on your desktop
echo.
echo    2. Click the Start button and type "GQM Lab Manager", 
echo       then click on the application icon
echo.
echo    3. Press the Windows key, type "GQM Lab Manager", 
echo       then click on the application icon
echo.
echo  Thank you for choosing GQM Lab Manager!
echo.
echo ================================================================================================
echo.
echo  Press any key to exit the installer...
pause >nul