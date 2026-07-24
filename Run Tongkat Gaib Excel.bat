@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "PYTHON_EXE="
set "PYTHON_LAUNCHER="

if exist ".venv\Scripts\python.exe" (
    set "PYTHON_EXE=.venv\Scripts\python.exe"
    "%PYTHON_EXE%" -c "import sys; raise SystemExit(0 if (sys.version_info >= (3, 10) and getattr(sys, '_is_gil_enabled', lambda: True)()) else 1)" >nul 2>&1
    if errorlevel 1 (
        echo.
        echo This project's .venv was created with an unsupported Python build
        echo ^(too old, or the experimental free-threaded Python^).
        echo Delete the .venv folder, install Python 3.12 ^(64-bit, standard^), then run again.
        echo.
        pause
        exit /b 1
    )
    goto :check_packages
)

echo First-time setup: creating the Python environment...

where py >nul 2>&1
if errorlevel 1 goto :no_python

rem Prefer stable Python versions with pre-built pandas wheels on Windows.
for %%V in (3.12 3.11 3.13 3.10) do (
    py -%%V -c "import sys; raise SystemExit(0 if (sys.version_info >= (3, 10) and getattr(sys, '_is_gil_enabled', lambda: True)()) else 1)" >nul 2>&1
    if not errorlevel 1 (
        set "PYTHON_LAUNCHER=py -%%V"
        goto :create_venv
    )
)

rem Last resort: whatever "py -3" points to (may lack wheels).
py -3 -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>&1
if errorlevel 1 goto :no_python
set "PYTHON_LAUNCHER=py -3"

:create_venv
echo Using %PYTHON_LAUNCHER% ...
%PYTHON_LAUNCHER% -m venv .venv
if errorlevel 1 goto :error

set "PYTHON_EXE=.venv\Scripts\python.exe"

:check_packages
"%PYTHON_EXE%" -c "import pandas, openpyxl, pydantic, typer" >nul 2>&1
if errorlevel 1 (
    echo Installing required packages...
    "%PYTHON_EXE%" -m pip install --upgrade pip setuptools wheel
    if errorlevel 1 goto :error
    "%PYTHON_EXE%" -m pip install --only-binary=pandas -r requirements.txt
    if errorlevel 1 goto :install_failed
)

"%PYTHON_EXE%" main.py
if errorlevel 1 goto :error
exit /b 0

:no_python
echo.
echo Python 3.10+ is required and was not found.
echo.
echo Install Python 3.12 from https://www.python.org/downloads/
echo During setup, check "Add python.exe to PATH".
echo Avoid the "free-threaded" experimental build.
echo.
pause
exit /b 1

:install_failed
echo.
echo Package install failed. This usually means your Python build has no
echo pre-built pandas wheel, so pip tried to compile from source.
echo.
"%PYTHON_EXE%" -c "import sys, platform; print('Detected Python:', sys.version); print('Executable:', sys.executable); print('Architecture:', platform.machine())"
echo.
echo Fix options:
echo   1. Install Python 3.12 ^(64-bit, standard build^) from python.org
echo   2. Delete the .venv folder in this project, then run this file again
echo   3. Ask for TongkatGaibExcel.exe ^(no Python needed^)
echo.
pause
exit /b 1

:error
echo.
echo The app could not start. Review the error above.
pause
exit /b 1
