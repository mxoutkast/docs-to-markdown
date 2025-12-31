@echo off
setlocal enabledelayedexpansion

echo ========================================
echo docs-to-markdown Launcher
echo ========================================
echo.

REM Check if virtual environment exists
if not exist ".venv" (
    echo ERROR: Virtual environment not found.
    echo Please run setup.bat first to create the virtual environment.
    echo.
    pause
    exit /b 1
)

REM Activate virtual environment
call .venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo ERROR: Failed to activate virtual environment.
    echo.
    pause
    exit /b 1
)

REM Run docs-to-markdown with all arguments
docs-to-markdown %*

REM Check exit code
if %errorlevel% neq 0 (
    echo.
    echo Application exited with error code %errorlevel%.
    echo.
    pause
    exit /b %errorlevel%
)

echo.
pause
