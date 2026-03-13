@echo off
setlocal

echo.
echo CTFMS EZ Certificates - setup and run
echo --------------------------------------

REM 1) Check that Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo Python 3 is not installed or not on PATH.
    echo.
    echo Please install Python 3 from:
    echo   https://www.python.org/downloads/windows/
    echo (Tick "Add python.exe to PATH" during installation.)
    echo.
    pause
    goto :eof
)

REM 2) Create a virtual environment (if it does not exist yet)
if not exist ".venv" (
    echo.
    echo Creating virtual environment (.venv)...
    python -m venv .venv
    if errorlevel 1 (
        echo.
        echo Failed to create virtual environment.
        pause
        goto :eof
    )
)

REM 3) Activate the virtual environment
call ".venv\Scripts\activate.bat"
if errorlevel 1 (
    echo.
    echo Failed to activate virtual environment.
    pause
    goto :eof
)

REM 4) Upgrade pip and install dependencies
echo.
echo Installing/Updating Python packages...
python -m pip install --upgrade pip
pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo Failed to install required Python packages.
    pause
    goto :eof
)

REM 5) Run the GUI application
echo.
echo Starting CTFMS EZ Certificates...
python certificate_app.py

echo.
echo Application closed.
pause

endlocal
