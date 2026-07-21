@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo First-time setup: creating the Python environment...
    where py >nul 2>&1
    if errorlevel 1 (
        echo.
        echo Python is not installed. Install Python 3.11 or newer, then run this file again.
        pause
        exit /b 1
    )

    py -3 -m venv .venv
    if errorlevel 1 goto :error
)

".venv\Scripts\python.exe" -c "import pandas, openpyxl, pydantic, typer" >nul 2>&1
if errorlevel 1 (
    echo Installing required packages...
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    if errorlevel 1 goto :error
)

".venv\Scripts\python.exe" main.py
if errorlevel 1 goto :error
exit /b 0

:error
echo.
echo The app could not start. Review the error above.
pause
exit /b 1
