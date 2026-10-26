@echo off
setlocal

REM ── Try to find Python ──────────────────────────────────────────────────────
set PYTHON=
set PIP=

REM 1. Windows Python Launcher (most reliable on Windows)
where py >nul 2>&1
if %errorlevel%==0 (
    set PYTHON=py
    set PIP=py -m pip
    goto :found
)

REM 2. python3
where python3 >nul 2>&1
if %errorlevel%==0 (
    set PYTHON=python3
    set PIP=python3 -m pip
    goto :found
)

REM 3. python
where python >nul 2>&1
if %errorlevel%==0 (
    set PYTHON=python
    set PIP=python -m pip
    goto :found
)

REM ── Python not found ────────────────────────────────────────────────────────
echo.
echo  ERROR: Python was not found on this computer.
echo.
echo  Please install Python from https://www.python.org/downloads/
echo  During installation, check the box that says:
echo    "Add Python to PATH"
echo.
echo  After installing, run this script again.
echo.
pause
exit /b 1

:found
echo  Using Python: %PYTHON%
echo.

REM ── Install / upgrade dependencies ─────────────────────────────────────────
echo  Installing dependencies (this only takes a moment)...
%PIP% install --upgrade pandas openpyxl scipy matplotlib
if %errorlevel% neq 0 (
    echo.
    echo  ERROR: pip install failed. See message above.
    pause
    exit /b 1
)

echo.
echo  Launching Growth Curve Analyzer...
echo.

REM ── Run the app ────────────────────────────────────────────────────────────
%PYTHON% growth_curve_analyzer.py

pause
