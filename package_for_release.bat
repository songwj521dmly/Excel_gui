@echo off
setlocal enabledelayedexpansion

echo ==================================================
echo   Excel Processor Packaging Script
echo ==================================================

set "QT_BIN_DIR=D:\anaconda3\Library\bin"
set "SOURCE_EXE=build\bin\excel_gui.exe"
set "DIST_DIR=dist"

if not exist "%QT_BIN_DIR%\windeployqt.exe" (
    echo [ERROR] windeployqt.exe not found at %QT_BIN_DIR%
    pause
    exit /b 1
)

if not exist "%SOURCE_EXE%" (
    echo [ERROR] %SOURCE_EXE% not found. Please build the project first using build_with_qt.bat.
    pause
    exit /b 1
)

:: Sanitize PATH to include only Qt bin and system paths
:: This prevents windeployqt from finding conflicting Qt versions in other paths
set "PATH=%QT_BIN_DIR%;%SystemRoot%\System32;%SystemRoot%;%SystemRoot%\System32\Wbem"

echo [INFO] Cleaning distribution directory...
if exist "%DIST_DIR%" rmdir /s /q "%DIST_DIR%"
mkdir "%DIST_DIR%"

echo [INFO] Copying executable...
copy "%SOURCE_EXE%" "%DIST_DIR%\" >nul
copy "build\bin\excel_console.exe" "%DIST_DIR%\" >nul 2>&1

echo [INFO] Running windeployqt to deploy dependencies...
echo This may take a few seconds...
"%QT_BIN_DIR%\windeployqt.exe" --release --no-translations --compiler-runtime "%DIST_DIR%\excel_gui.exe"

echo.
echo ==================================================
echo [SUCCESS] Packaging complete!
echo The application is ready in the "%DIST_DIR%" folder.
echo You can zip this folder and send it to other computers.
echo ==================================================
pause
