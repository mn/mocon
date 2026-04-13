@echo off
:: ============================================================
:: MOCON Automatic Balance Analysis System — AdvancedRx
:: Double-click this file to start the converter.
:: ============================================================

:: Fix for UNC path issue — CMD can't run from a network path
:: so we force it to a local drive first
if "%~d0"=="\\" (
    cd /d %USERPROFILE%
)

echo =============================================================
echo            MOCON Automatic Balance Analysis System
echo                        AdvancedRx
echo =============================================================

:: Verify Python is installed
echo Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH.
    echo Please install Python from https://python.org and try again.
    pause
    exit /b
)

:: Install required packages if missing
echo Checking required packages...
python -c "import reportlab, watchdog" >nul 2>&1
if errorlevel 1 (
    echo [INFO] Installing required packages...
    pip install reportlab watchdog
    if errorlevel 1 (
        echo [ERROR] Failed to install packages.
        echo Try running this bat file as Administrator.
        pause
        exit /b
    )
)

:: Create the Converter folder on the Desktop if it doesn't exist
if not exist "%USERPROFILE%\Desktop\Converter" (
    mkdir "%USERPROFILE%\Desktop\Converter"
    echo [INFO] Created Converter folder at %USERPROFILE%\Desktop\Converter
)

:: Run from a known local directory so relative paths work
cd /d "%USERPROFILE%\Desktop\RUN FOR MOCON" 2>nul || cd /d "%~dp0"

:: Start the converter
echo Starting the directory watcher...
echo Drop .PRN files into: %USERPROFILE%\Desktop\Converter
echo =============================================================
python mocon_converter.py
if errorlevel 1 (
    echo [ERROR] The script stopped unexpectedly.
    echo Check the error above and contact your administrator.
    pause
    exit /b
)

echo =============================================================
echo Watcher stopped. Press any key to exit.
echo =============================================================
pause
exit /b
