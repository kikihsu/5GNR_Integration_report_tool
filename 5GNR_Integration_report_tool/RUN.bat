@echo off
title KPI Log File Processing Tool

:: ── Change to the folder where this .bat file lives ──────────────────────────
:: This ensures relative paths (Logs/, KPI_template.xlsx) resolve correctly
:: regardless of where the user launches the file from.
cd /d "%~dp0"

:: ── Check Python is available ─────────────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo   [X] Python not found.
    echo       Please run INSTALL.bat first, or contact your system administrator.
    echo.
    pause
    exit /b 1
)

:: ── Check main.py exists ──────────────────────────────────────────────────────
if not exist "%~dp0main.py" (
    echo.
    echo   [X] main.py not found.
    echo       Please make sure all files from the package are in the same folder.
    echo.
    pause
    exit /b 1
)

:: ── Launch the tool ───────────────────────────────────────────────────────────
python main.py

:: ── Handle exit codes ─────────────────────────────────────────────────────────
if errorlevel 1 (
    echo.
    echo ========================================================================
    echo   [X] The tool exited with an error.
    echo       Scroll up to read the error message.
    echo       Take a screenshot and contact the tool owner if you need help.
    echo ========================================================================
    echo.
    pause
)
