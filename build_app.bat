@echo off
REM Build script for DataLens
REM This script packages the application into a standalone Windows executable

echo ========================================
echo DataLens - Build Script
echo ========================================
echo.

REM Clean previous builds
echo [1/4] Cleaning previous builds...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
echo Done.
echo.

REM Verify requirements are installed
echo [2/4] Verifying required packages...
pip show pandas >nul 2>&1
if errorlevel 1 (
    echo ERROR: pandas not installed. Installing from requirements.txt...
    pip install -r requirements.txt
)
pip show openpyxl >nul 2>&1
if errorlevel 1 (
    echo ERROR: openpyxl not installed. Installing from requirements.txt...
    pip install -r requirements.txt
)
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo ERROR: PyInstaller not installed. Installing...
    pip install pyinstaller
)
echo All required packages are installed.
echo.

REM Build the executable
echo [3/4] Building executable...
pyinstaller datalens.spec --clean
if errorlevel 1 (
    echo.
    echo ERROR: Build failed!
    pause
    exit /b 1
)
echo Done.
echo.

REM Display results
echo [4/4] Build completed successfully!
echo.
echo ========================================
echo Your executable is ready!
echo Location: dist\DataLens.exe
echo ========================================
echo.

REM Open the dist folder
if exist dist\DataLens.exe (
    echo Opening dist folder...
    explorer dist
) else (
    echo WARNING: Executable not found in dist folder!
)

echo.
pause
