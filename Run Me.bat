@echo off
REM Eclipse Swiss Army Knife - Batch Wrapper
REM This batch file launches the PowerShell installer with proper elevation

echo.
echo ========================================
echo  Eclipse Swiss Army Knife
echo ========================================
echo.
echo IMPORTANT: Before installing Eclipse Swiss Army Knife services, ensure that:
echo   1. The Eclipse Export tool has been run at least once
echo   2. Database connection settings have been configured
echo   3. The application works correctly when run manually
echo.

REM Check if launcher script exists
if not exist "%~dp0Launch-EclipseSwissArmyKnife.ps1" (
    echo ERROR: Launch-EclipseSwissArmyKnife.ps1 not found!
    echo Please ensure the launcher script is in the same directory.
    pause
    exit /b 1
)

REM Check if running as administrator
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo This installer requires administrator privileges.
    echo.
    echo Attempting to restart with administrator privileges...
    echo.
    
    REM Use PowerShell to elevate privileges (more reliable)
    powershell.exe -Command "Start-Process '%~f0' -Verb RunAs"
    
    REM Exit the current instance
    exit /b 0
)

REM Set execution policy for current process and run PowerShell launcher
echo Starting Eclipse Swiss Army Knife Launcher...
echo.
echo The launcher will check for updates and run the latest version.
echo.

powershell.exe -ExecutionPolicy Bypass -NoExit -File "%~dp0Launch-EclipseSwissArmyKnife.ps1" %*

if %errorLevel% neq 0 (
    echo.
    echo ERROR: PowerShell script failed with error code %errorLevel%
    echo.
    echo Possible solutions:
    echo 1. Make sure Launch-EclipseSwissArmyKnife.ps1 is in the same directory
    echo 2. Try running PowerShell as administrator and execute the launcher directly
    echo 3. Check if PowerShell execution policy is blocking the script
    echo 4. Ensure you have internet connectivity for update checks
    echo.
    pause
    exit /b 1
)

echo.
echo PowerShell script completed. The interactive menu should remain open.
echo If the menu closed unexpectedly, you can run the PowerShell script directly.
pause >nul 