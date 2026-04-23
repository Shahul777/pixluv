@echo off
title Basa Web App
cd /d "%~dp0"

echo ==================================================
echo Basa - Pre-Printing Workflow (Web App)
echo ==================================================
echo.

:: Check if Python is installed
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not in PATH.
    echo Please install Python from https://www.python.org/downloads/
    pause
    exit /b 1
)

:: Create virtual environment if it doesn't exist
if not exist ".venv" (
    echo [1/3] Creating virtual environment...
    python -m venv .venv
    if %errorlevel% neq 0 (
        echo [ERROR] Failed to create virtual environment.
        pause
        exit /b 1
    )
) else (
    echo [1/3] Virtual environment found.
)

:: Activate virtual environment and install dependencies
echo [2/3] Installing dependencies...
call .venv\Scripts\activate.bat
pip install -r requirements.txt --quiet
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b 1
)

:: Start the web app
echo [3/3] Starting Basa Web App...
echo.
python app.py
pause
