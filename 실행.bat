@echo off
chcp 65001 >nul
cd /d "%~dp0"
python launcher.py
if errorlevel 1 (
    echo.
    echo 오류가 발생했습니다. 위 내용을 확인해주세요.
    pause
)
