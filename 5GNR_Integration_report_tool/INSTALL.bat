@echo off
title KPI Tool - Environment Setup

echo.
echo ========================================================================
echo   KPI LOG FILE PROCESSING TOOL - Environment Setup
echo ========================================================================
echo.

:: ── Check Python is available ────────────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo   [X] Python not found.
    echo.
    echo       This tool requires Python to be installed by your IT department.
    echo       Please contact your system administrator before continuing.
    echo.
    pause
    exit /b 1
)

for /f "tokens=*" %%v in ('python --version 2^>^&1') do set PY_VER=%%v
echo   [OK] Found %PY_VER%
echo.

:: ── Check requirements.txt exists ────────────────────────────────────────────
if not exist "%~dp0requirements.txt" (
    echo   [X] requirements.txt not found.
    echo       Please make sure all files from the package are in the same folder.
    echo.
    pause
    exit /b 1
)

:: ── Install libraries ─────────────────────────────────────────────────────────
echo   Installing required libraries...
echo   (This may take a minute on first run)
echo.

pip install -r "%~dp0requirements.txt" --quiet
if errorlevel 1 (
    echo.
    echo   [X] Installation failed.
    echo       Please take a screenshot of this window and contact the tool owner.
    echo.
    pause
    exit /b 1
)

echo.
echo ========================================================================
echo   [OK] Installation complete! You can close this window.
echo        Next step: double-click RUN.bat to start the tool.
echo ========================================================================
echo.
pause
