@echo off
setlocal enabledelayedexpansion

echo ==================================================
echo   Excel Processor Release Packaging Script
echo ==================================================

:: 1. Clean build directory to ensure correct Qt version detection
if exist "build" (
    echo [INFO] Cleaning build directory...
    rmdir /s /q "build"
)

:: 2. Run the build script
call build_with_qt.bat
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Build failed. Aborting packaging.
    exit /b 1
)

set "BUILD_DIR=build"
set "DIST_DIR=dist"
set "EXE_NAME=excel_gui.exe"

:: 3. Find the executable
if exist "%BUILD_DIR%\bin\%EXE_NAME%" (
    set "EXE_PATH=%BUILD_DIR%\bin\%EXE_NAME%"
) else if exist "%BUILD_DIR%\Release\%EXE_NAME%" (
    set "EXE_PATH=%BUILD_DIR%\Release\%EXE_NAME%"
) else (
    echo [ERROR] Could not find %EXE_NAME% in %BUILD_DIR%\bin or %BUILD_DIR%\Release
    exit /b 1
)

echo [INFO] Found executable at: %EXE_PATH%

:: 4. Define Qt Path (Qt 5.12.12)
set "QT_PATH=D:\Qt\Qt5.12.12\5.12.12\msvc2017_64"
set "WINDEPLOYQT=%QT_PATH%\bin\windeployqt.exe"

if not exist "%WINDEPLOYQT%" (
    echo [ERROR] windeployqt.exe not found at %WINDEPLOYQT%
    exit /b 1
)

:: 5. Prepare distribution directory
if exist "%DIST_DIR%" rmdir /s /q "%DIST_DIR%"
mkdir "%DIST_DIR%"

echo [INFO] Copying executable to %DIST_DIR%...
copy "%EXE_PATH%" "%DIST_DIR%\" >nul

:: 6. Run windeployqt
echo [INFO] Running windeployqt to deploy dependencies...
echo [CMD] "%WINDEPLOYQT%" --release --compiler-runtime --no-translations --dir "%DIST_DIR%" "%DIST_DIR%\%EXE_NAME%"

"%WINDEPLOYQT%" --release --compiler-runtime --no-translations --dir "%DIST_DIR%" "%DIST_DIR%\%EXE_NAME%"

if %ERRORLEVEL% neq 0 (
    echo [ERROR] windeployqt failed.
    exit /b 1
)

:: 7. Deploy MSVC Runtime (CRITICAL for stranger's computer)
echo [INFO] Deploying MSVC Runtime...

:: Path to MSVC Redist CRT folder on this specific machine (VS 2026 Preview)
set "MSVC_CRT_DIR=D:\Program Files\Microsoft Visual Studio\18\Insiders\VC\Redist\MSVC\14.50.35710\x64\Microsoft.VC145.CRT"

if exist "%MSVC_CRT_DIR%" (
    echo [INFO] Found MSVC CRT at: %MSVC_CRT_DIR%
    echo [INFO] Copying runtime DLLs...
    copy "%MSVC_CRT_DIR%\*.dll" "%DIST_DIR%\" >nul
) else (
    echo [WARNING] Could not find MSVC CRT at expected path.
    echo [WARNING] The application might not run on computers without VS installed.
    echo [WARNING] Expected: %MSVC_CRT_DIR%
    
    :: Try to fallback to finding any CRT folder (basic attempt)
    :: But for now, warning is enough as we know the path exists.
)

echo.
echo ==================================================
echo   Packaging Complete!
echo   Output directory: %CD%\%DIST_DIR%
echo ==================================================
echo.
echo [IMPORTANT]
echo To share this program:
echo 1. Zip the entire '%DIST_DIR%' folder.
echo 2. Send the zip file to the other computer.
echo 3. Tell them to unzip and run 'excel_gui.exe'.
echo.
