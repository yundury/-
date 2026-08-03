@echo off
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python was not found.
    echo Please install Python from https://www.python.org/downloads/
    echo During installation, make sure to check "Add python.exe to PATH".
    pause
    exit /b 1
)

if not exist "%~dp0.setup_done" (
    echo Setting up for the first time. This can take a few minutes, please wait...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo [ERROR] Package install failed. Check your internet connection and try again.
        pause
        exit /b 1
    )
    python -m playwright install chromium
    if errorlevel 1 (
        echo [ERROR] Browser install failed. Check your internet connection and try again.
        pause
        exit /b 1
    )
    echo done> "%~dp0.setup_done"
    echo Setup complete!
)

start "" pythonw launcher.py
