@echo off
REM ============================================================
REM Build script for Lender Feedback Tool - Windows executable
REM ============================================================
REM
REM Prerequisites:
REM   1. Install Python 3.10+ from https://python.org (one-time)
REM   2. Run this script from the project folder
REM
REM Output: dist\LenderFeedbackTool.exe (single file, ~40-50 MB)
REM ============================================================

echo.
echo ============================================================
echo Lender Feedback Tool - Windows Build
echo ============================================================
echo.

REM Check Python is installed
where python >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Please install Python 3.10+ from https://python.org
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

echo [1/4] Creating virtual environment...
if not exist venv (
    python -m venv venv
)

echo [2/4] Installing dependencies...
call venv\Scripts\activate.bat
python -m pip install --upgrade pip --quiet
python -m pip install flask pdfplumber python-docx pyinstaller --quiet

echo [3/4] Building executable with PyInstaller...
echo (This takes 2-5 minutes on first run)
pyinstaller build.spec --clean --noconfirm

echo [4/4] Done!
echo.
if exist dist\LenderFeedbackTool.exe (
    echo ============================================================
    echo SUCCESS: Built dist\LenderFeedbackTool.exe
    echo ============================================================
    echo.
    echo You can now:
    echo   - Double-click dist\LenderFeedbackTool.exe to run the tool
    echo   - Copy this single .exe to any Windows computer ^(no Python needed^)
    echo   - Share it with colleagues ^(self-contained^)
    echo.
) else (
    echo ============================================================
    echo BUILD FAILED - check errors above
    echo ============================================================
)
pause
