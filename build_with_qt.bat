@echo off
setlocal enabledelayedexpansion

echo ==================================================
echo   Excel Processor Build Script (Qt 5.12.12)
echo ==================================================

:: 1. Configure path variables
set "QT_PATH=D:\Qt\Qt5.12.12\5.12.12\msvc2017_64"
set "QT_TOOLS_CMAKE=D:\Qt\Qt5.12.12\Tools\CMake_64\bin"
set "QT_TOOLS_NINJA=D:\Qt\Qt5.12.12\Tools\Ninja"
set "BUILD_DIR=build"

:: Check if internal CMake/Ninja exists
if not exist "%QT_TOOLS_CMAKE%" set "QT_TOOLS_CMAKE="
if not exist "%QT_TOOLS_NINJA%" set "QT_TOOLS_NINJA="

:: 2. Check Qt path
if not exist "%QT_PATH%" (
    echo [ERROR] Qt path not found at: %QT_PATH%
    exit /b 1
)

:: 3. Add Qt tools to PATH
if not "%QT_TOOLS_CMAKE%"=="" set "PATH=%QT_TOOLS_CMAKE%;%PATH%"
if not "%QT_TOOLS_NINJA%"=="" set "PATH=%QT_TOOLS_NINJA%;%PATH%"

:: 4. Setup MSVC
where cl >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo [INFO] Loading MSVC environment...
    set "VCVARS_PATH="
    if exist "%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe" (
        for /f "usebackq tokens=*" %%i in (`"%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe" -latest -prerelease -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath`) do (
            if exist "%%i\VC\Auxiliary\Build\vcvars64.bat" set "VCVARS_PATH=%%i\VC\Auxiliary\Build\vcvars64.bat"
        )
    )
    if "!VCVARS_PATH!"=="" (
         echo [ERROR] vcvars64.bat not found.
         exit /b 1
    )
    call "!VCVARS_PATH!"
)

:: 5. Check CMake
where cmake >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo [ERROR] CMake not found.
    exit /b 1
)

:: 6. Configure
if not exist "%BUILD_DIR%" mkdir "%BUILD_DIR%"
cd "%BUILD_DIR%"

echo [INFO] Configuring CMake...
set "GENERATOR=Ninja"
where ninja >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo [WARNING] Ninja not found. Using NMake.
    set "GENERATOR=NMake Makefiles"
)

echo [INFO] Using Generator: %GENERATOR%

cmake -G "%GENERATOR%" -DCMAKE_BUILD_TYPE=Release -DCMAKE_PREFIX_PATH="%QT_PATH%" -DCMAKE_TOOLCHAIN_FILE="%QT_PATH%\lib\cmake\Qt5\Qt5Config.cmake" -DOPTION_FORCE_QT5=ON ..

if %ERRORLEVEL% neq 0 (
    echo [ERROR] CMake configuration failed.
    cd ..
    exit /b 1
)

:: 7. Build
echo [INFO] Building project...
cmake --build . --config Release
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Build failed.
    cd ..
    exit /b 1
)

echo [INFO] Build successful!
cd ..
