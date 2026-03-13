@echo off
cd /d "%~dp0"
setlocal EnableExtensions
title Document to Markdown Converter
echo ============================================================
echo   Document to Markdown Converter
echo   Converts all PDF and DOCX files in THIS folder
echo   Output goes to: md_output\
echo ============================================================
echo.

REM ===== Step 1: Check if Python is already installed =====
python --version >nul 2>&1
if errorlevel 1 (
    echo Python not found. Attempting automatic install...
    echo.

    REM Check if winget is available on this machine
    winget --version >nul 2>&1
    if errorlevel 1 (
        echo ERROR: winget is not available on this machine.
        echo Please install Python manually from https://www.python.org/downloads/
        echo Make sure to check "Add Python to PATH" during install, then re-run this file.
        pause
        exit /b 1
    )

    echo Installing Python 3.11 via winget - this may take a few minutes...
    winget install Python.Python.3.11 --silent --accept-package-agreements --accept-source-agreements
    if errorlevel 1 (
        echo.
        echo ERROR: Automatic Python install failed.
        echo Please install Python manually from https://www.python.org/downloads/
        echo Make sure to check "Add Python to PATH" during install, then re-run this file.
        pause
        exit /b 1
    )

    echo.
    echo Python installed successfully!
    echo Refreshing environment...

    REM Refresh PATH so python is visible in this session
    for /f "tokens=*" %%i in ('where python 2^>nul') do set PYTHON_EXE=%%i
    if "%PYTHON_EXE%"=="" (
        echo.
        echo Python was installed but PATH was not updated yet.
        echo Please CLOSE this window and double-click CONVERT_DOCS.bat again.
        pause
        exit /b 0
    )
)

echo Python found. OK
echo.

REM ===== Step 2: Install Python libraries =====
echo Checking / installing required libraries...
python -m pip install --upgrade pip --quiet --no-warn-script-location
python -m pip install -r requirements_extract.txt --quiet --no-warn-script-location
if errorlevel 1 (
    echo.
    echo ERROR: Failed to install required libraries.
    echo Try right-clicking CONVERT_DOCS.bat and choosing "Run as Administrator".
    pause
    exit /b 1
)

echo Libraries ready. OK
echo.

REM ===== Step 3: Run the converter =====
echo Starting conversion...
echo.
python pdf_docx_to_markdown.py || (
    echo.
    echo ERROR: Conversion failed. See message above.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   Finished! Press any key to open the output folder...
echo ============================================================
pause
if exist "md_output" explorer "md_output"
