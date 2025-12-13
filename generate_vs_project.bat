@echo off
setlocal enabledelayedexpansion

echo ==================================================
echo   Generate Visual Studio Solution
echo ==================================================

:: 1. Configuration
set "QT_PATH=D:\Qt\6.9.1\msvc2022_64"
set "BUILD_DIR=build_vs"

:: 2. Setup Tools
set "QT_TOOLS_CMAKE=D:\Qt\Tools\CMake_64\bin"
set "PATH=%QT_TOOLS_CMAKE%;%PATH%"

:: 3. Check for CMake
where cmake >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo [ERROR] CMake not found. Please ensure CMake is installed.
    pause
    exit /b 1
)

:: 4. Create Build Directory
if not exist "%BUILD_DIR%" mkdir "%BUILD_DIR%"
cd "%BUILD_DIR%"

:: 5. Generate Solution
echo [INFO] Generating Visual Studio Solution...
echo [INFO] Using Qt Path: %QT_PATH%

:: We use "Visual Studio 17 2022" generator. 
:: VS 2026 (v18) is backward compatible and will open this solution without issues.
cmake -G "Visual Studio 17 2022" -A x64 -DCMAKE_PREFIX_PATH="%QT_PATH%" ..

if %ERRORLEVEL% equ 0 (
    echo.
    echo ==================================================
    echo   [SUCCESS] Solution generated!
    echo ==================================================
    echo.
    echo You can now open the project in Visual Studio 2026:
    echo File: %CD%\ExcelProcessor.sln
    echo.
    echo Press any key to open the solution automatically...
    pause >nul
    start ExcelProcessor.sln
) else (
    echo.
    echo [ERROR] Failed to generate solution.
    echo Please check the error messages above.
    pause
)
