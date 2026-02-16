@echo off
REM Build the 32-bit SoftVoice host executable using PyInstaller.
REM Requires: py -3.14-32 (32-bit Python 3.14 launcher)
REM Requires: pip install pyinstaller (in the 32-bit Python)

cd /d "%~dp0"

echo Building 32-bit host executable...
py -3.14-32 -m PyInstaller --onefile --noconsole --name softvoice_host32 ^
    --add-data "synthDrivers\_ipc.py;." ^
    synthDrivers\host_softvoice32.py

if errorlevel 1 (
    echo Build failed!
    exit /b 1
)

echo Copying host executable to synthDrivers...
copy /y dist\softvoice_host32.exe synthDrivers\softvoice_host32.exe

echo Done. Files ready in synthDrivers\
echo.
echo To package as .nvda-addon, run: python build.py
