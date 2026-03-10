@echo off
:: ============================================================
:: eSTAR Package Extractor — Launcher
:: ============================================================
:: Double-click this file to launch the eSTAR extractor GUI.
:: Python 3.8+ must be installed and on PATH.
:: All pip dependencies are installed automatically on first run.
:: ============================================================

title eSTAR Package Extractor

:: Check Python is available
where python >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo.
    echo  ERROR: Python was not found on your system PATH.
    echo.
    echo  Please install Python 3.8 or later from https://www.python.org
    echo  and make sure "Add Python to PATH" is checked during installation.
    echo.
    pause
    exit /b 1
)

:: Ensure pip is available
python -m pip --version >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo.
    echo  ERROR: pip is not available.  Attempting to bootstrap it...
    python -m ensurepip --default-pip
)

:: Run the extractor script (located in the same folder as this .bat)
python "%~dp0estar_extractor.py"

:: If the script exits with an error, keep the window open
if %ERRORLEVEL% neq 0 (
    echo.
    echo  The script exited with an error.  See messages above.
    pause
)
