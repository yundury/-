@echo off
cd /d "%~dp0"
python launcher.py
if errorlevel 1 (
    echo.
    echo An error occurred. Please check the messages above.
    pause
)
