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

set "PYTHON_EXE="

REM Prefer the py launcher, then python on PATH. Require 3.10+.
where py >nul 2>&1
if not errorlevel 1 (
    py -3 -c "import sys; raise SystemExit(0 if sys.version_info >= (3,10) else 1)" >nul 2>&1
    if not errorlevel 1 (
        for /f "delims=" %%i in ('py -3 -c "import sys; print(sys.executable)"') do set "PYTHON_EXE=%%i"
    )
)

if not defined PYTHON_EXE (
    where python >nul 2>&1
    if not errorlevel 1 (
        python -c "import sys; raise SystemExit(0 if sys.version_info >= (3,10) else 1)" >nul 2>&1
        if not errorlevel 1 (
            for /f "delims=" %%i in ('python -c "import sys; print(sys.executable)"') do set "PYTHON_EXE=%%i"
        )
    )
)

if not defined PYTHON_EXE (
    echo Python 3.10+ not found. Attempting automatic install...
    echo.

    winget --version >nul 2>&1
    if errorlevel 1 (
        echo ERROR: winget is not available on this machine.
        echo Please install Python 3.10 or newer from https://www.python.org/downloads/
        echo Make sure to check "Add Python to PATH" during install, then re-run this file.
        pause
        exit /b 1
    )

    echo Installing Python 3.11 via winget - this may take a few minutes...
    winget install Python.Python.3.11 --silent --accept-package-agreements --accept-source-agreements
    if errorlevel 1 (
        echo.
        echo ERROR: Automatic Python install failed.
        echo Please install Python 3.10 or newer from https://www.python.org/downloads/
        echo Make sure to check "Add Python to PATH" during install, then re-run this file.
        pause
        exit /b 1
    )

    if exist "%LocalAppData%\Programs\Python\Python311\python.exe" (
        set "PYTHON_EXE=%LocalAppData%\Programs\Python\Python311\python.exe"
    )
    if not defined PYTHON_EXE (
        echo.
        echo Python was installed but PATH was not updated yet.
        echo Please CLOSE this window and double-click CONVERT_DOCS.bat again.
        pause
        exit /b 0
    )
)

echo Python found. OK
echo.

if not exist ".venv\Scripts\python.exe" (
    echo Creating local virtual environment...
    "%PYTHON_EXE%" -m venv .venv
    if errorlevel 1 (
        echo ERROR: Failed to create a local virtual environment.
        pause
        exit /b 1
    )
)

set "VPY=.venv\Scripts\python.exe"

echo Checking / installing required libraries...
"%VPY%" -c "import fitz, pdfplumber, docx, tqdm" >nul 2>&1
if errorlevel 1 (
    "%VPY%" -m pip install -r requirements_extract.txt --quiet --no-warn-script-location
    if errorlevel 1 (
        echo.
        echo ERROR: Failed to install required libraries.
        echo Check your internet connection and try again.
        pause
        exit /b 1
    )
)

echo Libraries ready. OK
echo.

echo Starting conversion...
echo.
"%VPY%" pdf_docx_to_markdown.py -y
if errorlevel 1 (
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
