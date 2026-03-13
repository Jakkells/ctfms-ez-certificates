@echo off
setlocal

REM Always run from this script's folder
cd /d "%~dp0"

echo.
echo Building CTFMS EZ Certificates EXE
echo -----------------------------------

REM 1) Check that Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo Python 3 is not installed or not on PATH.
    echo Please install Python 3 from:
    echo   https://www.python.org/downloads/windows/
    echo.
    pause
    goto :eof
)

REM 2) Make sure PyInstaller is available
echo.
echo Installing/Updating PyInstaller...
python -m pip install --upgrade pip
python -m pip install pyinstaller
if errorlevel 1 (
    echo.
    echo Failed to install PyInstaller.
    pause
    goto :eof
)

REM 3) Run PyInstaller to create a single EXE
REM --onefile   : pack everything into one executable
REM --noconsole : hide console window (GUI app)
REM --name      : exe name
REM --add-data  : include the Word template in the same directory as the exe
echo.
echo Running PyInstaller...
pyinstaller ^
  --onefile ^
  --noconsole ^
  --name "CTFMS-EZ-Certificates" ^
  --add-data "SERVICE CERTIFICATE.docx;." ^
  certificate_app.py

if errorlevel 1 (
    echo.
    echo PyInstaller build failed.
    pause
    goto :eof
)

echo.
echo Build finished.
echo The EXE should be at:
echo   dist\CTFMS-EZ-Certificates.exe
echo.
echo You can upload that EXE to GitHub Releases or share it directly.
echo.
pause

endlocal
