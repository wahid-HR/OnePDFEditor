@echo off
REM ============================================================================
REM  Offline PDF Editor – Windows build script
REM  Run this from the project root on a Windows machine with Python installed.
REM ============================================================================

echo.
echo ============================================================
echo  Building Offline PDF Editor for Windows
echo ============================================================
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not in PATH. Please install Python 3.8+ and try again.
    pause
    exit /b 1
)

REM Optional: create / activate venv
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
)
call venv\Scripts\activate.bat

echo Installing / updating dependencies...
python -m pip install --upgrade pip
pip install -r requirements.txt

echo.
echo Running PyInstaller...
pyinstaller --noconfirm OfflinePDFEditor.spec

if errorlevel 1 (
    echo.
    echo BUILD FAILED.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  BUILD SUCCEEDED
echo ============================================================
echo.
echo Executable is located at:
echo   dist\OfflinePDFEditor.exe
echo.
echo You can copy this single .exe to any Windows PC.
echo No Python installation is required on the target machine.
echo The application works completely offline.
echo.
pause
