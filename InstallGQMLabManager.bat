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
    echo ERROR: Not running as administrator >> %LOG_FILE%
    echo Installation aborted. Please run as administrator. >> %LOG_FILE%
    echo Log file created at: %LOG_FILE%
    pause
    exit /b
)

echo Running as administrator: YES >> %LOG_FILE%

:: Create a temporary directory
set TEMP_DIR=%TEMP%\GQMLabInstall
mkdir "%TEMP_DIR%" 2>nul

echo Preparing for installation...
echo Temp directory: %TEMP_DIR% >> %LOG_FILE%

:: Debug info
echo Current directory: %CD% >> %LOG_FILE%
dir /b >> %LOG_FILE%

echo Installing certificate...
echo Certificate installation started >> %LOG_FILE%

:: First try with local cer file (check all possible names to handle case sensitivity)
set CERT_FOUND=0
if exist "GQMLab_Simple.cer" (
    echo Found: GQMLab_Simple.cer
    echo Found certificate: GQMLab_Simple.cer >> %LOG_FILE%
    copy "GQMLab_Simple.cer" "%TEMP_DIR%\cert.cer" >nul
    set CERT_FOUND=1
) else if exist "gqmlab_simple.cer" (
    echo Found: gqmlab_simple.cer
    echo Found certificate: gqmlab_simple.cer >> %LOG_FILE%
    copy "gqmlab_simple.cer" "%TEMP_DIR%\cert.cer" >nul
    set CERT_FOUND=1
) else if exist "GQMlab_simple.cer" (
    echo Found: GQMlab_simple.cer
    echo Found certificate: GQMlab_simple.cer >> %LOG_FILE%
    copy "GQMlab_simple.cer" "%TEMP_DIR%\cert.cer" >nul
    set CERT_FOUND=1
)

:: If not found locally, download it
if %CERT_FOUND%==0 (
    echo Certificate not found locally, downloading...
    echo Certificate not found locally, attempting download >> %LOG_FILE%
    powershell -Command "Invoke-WebRequest -Uri 'https://gqmlab-solutions.github.io/msix-updates/GQMLab_Simple.cer' -OutFile '%TEMP_DIR%\cert.cer'"
    if %errorLevel% neq 0 (
        echo Failed to download certificate from primary URL >> %LOG_FILE%
        echo Trying alternative URL...
        powershell -Command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/gqmlab-solutions/msix-updates/main/GQMLab_Simple.cer' -OutFile '%TEMP_DIR%\cert.cer'"
        if %errorLevel% neq 0 (
            echo All download attempts failed. >> %LOG_FILE%
            echo *** ERROR: Could not download certificate *** >> %LOG_FILE%
            echo All download attempts failed.
            echo Log file created at: %LOG_FILE%
            echo Please send this log file to support for assistance.
            pause
            exit /b
        ) else {
            echo Successfully downloaded certificate from alternative URL >> %LOG_FILE%
        }
    } else {
        echo Successfully downloaded certificate from primary URL >> %LOG_FILE%
    }
)

:: Install certificate to BOTH Trusted Root and Trusted People stores
echo Installing certificate to Trusted Root store...
certutil -addstore Root "%TEMP_DIR%\cert.cer" >> %LOG_FILE% 2>&1
if %errorLevel% neq 0 (
    echo Certificate installation to Root store failed. >> %LOG_FILE%
    echo *** ERROR: Failed to install certificate to Root store *** >> %LOG_FILE%
    echo Certificate installation to Root store failed.
    echo Log file created at: %LOG_FILE%
    echo Please send this log file to support for assistance.
    pause
    exit /b
)

echo Installing certificate to Trusted People store...
certutil -addstore TrustedPeople "%TEMP_DIR%\cert.cer" >> %LOG_FILE% 2>&1
if %errorLevel% neq 0 (
    echo Certificate installation to Trusted People store failed. >> %LOG_FILE%
    echo *** ERROR: Failed to install certificate to Trusted People store *** >> %LOG_FILE%
    echo Certificate installation to Trusted People store failed.
    echo Log file created at: %LOG_FILE%
    echo Please send this log file to support for assistance.
    pause
    exit /b
)

echo Certificate installed successfully.
echo Certificate installed successfully. >> %LOG_FILE%
echo.

echo Installing GQM Lab Manager...
echo MSIX installation started >> %LOG_FILE%

:: Check if MSIX exists locally
set MSIX_FOUND=0
if exist "gqm_lab_manager.msix" (
    echo Found: gqm_lab_manager.msix
    echo Found MSIX: gqm_lab_manager.msix >> %LOG_FILE%
    copy "gqm_lab_manager.msix" "%TEMP_DIR%\app.msix" >nul
    set MSIX_FOUND=1
) else if exist "GQM_Lab_Manager.msix" (
    echo Found: GQM_Lab_Manager.msix
    echo Found MSIX: GQM_Lab_Manager.msix >> %LOG_FILE%
    copy "GQM_Lab_Manager.msix" "%TEMP_DIR%\app.msix" >nul
    set MSIX_FOUND=1
)

:: If not found locally, download it
if %MSIX_FOUND%==0 (
    echo MSIX package not found locally, downloading...
    echo MSIX not found locally, attempting download >> %LOG_FILE%
    powershell -Command "Invoke-WebRequest -Uri 'https://gqmlab-solutions.github.io/msix-updates/gqm_lab_manager.msix' -OutFile '%TEMP_DIR%\app.msix'"
    if %errorLevel% neq 0 (
        echo Failed to download MSIX package from primary URL >> %LOG_FILE%
        echo Trying alternative URL...
        powershell -Command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/gqmlab-solutions/msix-updates/main/gqm_lab_manager.msix' -OutFile '%TEMP_DIR%\app.msix'"
        if %errorLevel% neq 0 (
            echo All download attempts failed. >> %LOG_FILE%
            echo *** ERROR: Could not download MSIX package *** >> %LOG_FILE%
            echo All download attempts failed.
            echo Log file created at: %LOG_FILE%
            echo Please send this log file to support for assistance.
            pause
            exit /b
        } else {
            echo Successfully downloaded MSIX from alternative URL >> %LOG_FILE%
        }
    } else {
        echo Successfully downloaded MSIX from primary URL >> %LOG_FILE%
    }
)

:: Add developer mode option to bypass some certificate issues
echo Enabling developer mode features...
echo Enabling developer mode >> %LOG_FILE%
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock" /t REG_DWORD /f /v "AllowDevelopmentWithoutDevLicense" /d "1" >> %LOG_FILE% 2>&1

:: Try multiple installation methods
echo Installing package (method 1)...
echo Attempting installation method 1 >> %LOG_FILE%
powershell -Command "Add-AppxPackage -Path '%TEMP_DIR%\app.msix' -ForceApplicationShutdown -ForceUpdateFromAnyVersion | Out-String" > "%TEMP_DIR%\install_output.txt" 2>&1
type "%TEMP_DIR%\install_output.txt" >> %LOG_FILE%

if %errorLevel% neq 0 (
    echo Method 1 failed, trying method 2... >> %LOG_FILE%
    echo Method 1 failed, trying method 2 (Development Mode)...
    echo Attempting installation method 2 (Development Mode) >> %LOG_FILE%
    powershell -Command "Add-AppxPackage -Path '%TEMP_DIR%\app.msix' -ForceApplicationShutdown -ForceUpdateFromAnyVersion -DevelopmentMode | Out-String" > "%TEMP_DIR%\install_output2.txt" 2>&1
    type "%TEMP_DIR%\install_output2.txt" >> %LOG_FILE%
    
    if %errorLevel% neq 0 (
        echo Method 2 failed, trying method 3... >> %LOG_FILE%
        echo Method 2 failed, trying method 3 (Certificate extraction)...
        echo Attempting installation method 3 (Certificate extraction) >> %LOG_FILE%
        
        :: Try with direct certutil import
        echo Importing certificate directly from the MSIX package...
        mkdir "%TEMP_DIR%\extract" 2>nul
        echo Extracting MSIX package >> %LOG_FILE%
        powershell -Command "Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::ExtractToDirectory('%TEMP_DIR%\app.msix', '%TEMP_DIR%\extract')" >> %LOG_FILE% 2>&1
        
        echo Searching for certificates in extracted package >> %LOG_FILE%
        :: Find and install any certificates in the package
        for /r "%TEMP_DIR%\extract" %%f in (*.cer) do (
            echo Found certificate: %%f >> %LOG_FILE%
            echo Found certificate: %%f
            certutil -addstore Root "%%f" >> %LOG_FILE% 2>&1
            certutil -addstore TrustedPeople "%%f" >> %LOG_FILE% 2>&1
        )
        
        :: Try install again
        echo Attempting final installation >> %LOG_FILE%
        powershell -Command "Add-AppxPackage -Path '%TEMP_DIR%\app.msix' -ForceUpdateFromAnyVersion | Out-String" > "%TEMP_DIR%\install_output3.txt" 2>&1
        type "%TEMP_DIR%\install_output3.txt" >> %LOG_FILE%
        
        if %errorLevel% neq 0 (
            echo All installation attempts failed. >> %LOG_FILE%
            echo *** ERROR: All installation methods failed *** >> %LOG_FILE%
            
            :: Get the app package identity for troubleshooting
            echo Checking if app is already installed >> %LOG_FILE%
            powershell -Command "Get-AppxPackage | Where-Object {$_.Name -like '*gqm*' -or $_.Name -like '*lab*'} | Format-Table Name,PackageFullName -AutoSize | Out-String" >> %LOG_FILE%
            
            echo.
            echo ===== INSTALLATION FAILED =====
            echo All installation attempts failed.
            echo.
            echo A detailed log has been saved to your Desktop:
            echo %LOG_FILE%
            echo.
            echo Please send this log file to support for assistance.
            echo.
            echo Technical details:
            type "%TEMP_DIR%\install_output3.txt"
            echo.
            pause
            exit /b
        }
    }
)

:: Check if app was installed successfully
echo Verifying installation >> %LOG_FILE%
powershell -Command "$app = Get-AppxPackage | Where-Object {$_.Name -like '*gqm*' -or $_.Name -like '*lab*'}; if ($app) { Write-Host 'App found: ' + $app.Name; exit 0 } else { Write-Host 'App not found after installation'; exit 1 }" >> %LOG_FILE% 2>&1

if %errorLevel% neq 0 (
    echo *** WARNING: App not found in installed packages *** >> %LOG_FILE%
    echo.
    echo Installation completed, but the app could not be verified.
    echo This may be due to naming differences.
) else (
    echo App verified in installed packages >> %LOG_FILE%
)

:: Get app ID and create a shortcut
echo Creating shortcuts >> %LOG_FILE%
powershell -Command "$app = Get-AppxPackage | Where-Object {$_.Name -like '*gqm*' -or $_.Name -like '*lab*'}; if ($app) { $appId = $app.PackageFamilyName; Write-Host 'App ID: ' + $appId; } else { Write-Host 'App not found'; exit 1 }" > "%TEMP_DIR%\app_id.txt" 2>&1
set /p APP_ID=<"%TEMP_DIR%\app_id.txt"
echo %APP_ID% >> %LOG_FILE%

:: Create desktop shortcut for convenience
echo Creating desktop shortcut >> %LOG_FILE%
powershell -Command "$appid='%APP_ID%'; $Shell = New-Object -ComObject WScript.Shell; $Shortcut = $Shell.CreateShortcut('%USERPROFILE%\Desktop\GQM Lab Manager.lnk'); $Shortcut.TargetPath = 'shell:AppsFolder\' + $appId; $Shortcut.Save()" >> %LOG_FILE% 2>&1

echo.
echo GQM Lab Manager installed successfully!
echo Installation completed successfully >> %LOG_FILE%
echo.
echo You can now launch the application by:
echo  1. Clicking the Start button
echo  2. Typing "GQM Lab Manager"
echo  3. Clicking on the GQM Lab Manager app icon
echo.
echo A shortcut has also been created on your Desktop.
echo.
echo If you can't find the app, check your Desktop for a shortcut
echo or check the installation log at:
echo %LOG_FILE%
echo.
echo Press any key to exit...
pause > nul

:: Clean up
echo Cleaning up temporary files >> %LOG_FILE%
rmdir /s /q "%TEMP_DIR%" >> %LOG_FILE% 2>&1
echo Installation process completed at: %DATE% %TIME% >> %LOG_FILE%