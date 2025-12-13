@echo off
setlocal

echo ==================================================
echo   Clean Environment Simulation Test
echo ==================================================
echo.
echo This script simulates a computer without Qt or VS installed.
echo It clears the PATH environment variable to standard Windows defaults.
echo.

:: 1. Check if dist exists
if not exist "dist\excel_gui.exe" (
    echo [ERROR] dist\excel_gui.exe not found.
    echo Please run package_release.bat first.
    pause
    exit /b 1
)

:: 2. Simulate Clean Environment
echo [INFO] Original PATH length: %PATH:~0,100%...
echo [INFO] Clearing PATH...

:: Set PATH to minimal Windows defaults
set "PATH=C:\Windows\System32;C:\Windows;C:\Windows\System32\Wbem"

echo [INFO] New PATH: %PATH%
echo.
echo [INFO] Launching application from 'dist' folder...
echo If the application starts successfully, the deployment is valid.
echo If you see "System Error" or missing DLLs, the deployment is incomplete.
echo.

cd dist
start "" "excel_gui.exe"

echo Application launched. Check if it opened correctly.
pause
