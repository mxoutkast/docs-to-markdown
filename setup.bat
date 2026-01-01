@echo off
setlocal enabledelayedexpansion

echo ========================================
echo docs-to-markdown Setup Script
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH.
    echo Please install Python 3.9 or higher from https://www.python.org/
    echo.
    pause
    exit /b 1
)

echo Checking Python version...
for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
echo Found Python %PYTHON_VERSION%
echo.

REM Check if .venv already exists
if exist ".venv" (
    echo Virtual environment already exists.
    set /p RECREATE="Do you want to recreate it? (y/N): "
    if /i "!RECREATE!"=="y" (
        echo Removing existing virtual environment...
        rmdir /s /q .venv
        if %errorlevel% neq 0 (
            echo ERROR: Failed to remove existing .venv directory.
            echo Please close any applications using the virtual environment and try again.
            echo.
            pause
            exit /b 1
        )
    ) else (
        echo Skipping virtual environment creation.
        goto :install_deps
    )
)

REM Create virtual environment
echo Creating virtual environment...
python -m venv .venv
if %errorlevel% neq 0 (
    echo ERROR: Failed to create virtual environment.
    echo Please check that you have Python 3.9+ installed and have sufficient permissions.
    echo.
    pause
    exit /b 1
)
echo Virtual environment created successfully.
echo.

:install_deps
REM Activate virtual environment and install dependencies
echo Installing dependencies...
call .venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo ERROR: Failed to activate virtual environment.
    echo.
    pause
    exit /b 1
)

python -m pip install --upgrade pip
if %errorlevel% neq 0 (
    echo ERROR: Failed to upgrade pip.
    echo.
    pause
    exit /b 1
)

python -m pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ERROR: Failed to install dependencies from requirements.txt.
    echo Please check your internet connection and try again.
    echo.
    pause
    exit /b 1
)
echo Dependencies installed successfully.
echo.

REM Install package in editable mode
echo Installing docs-to-markdown package...
python -m pip install -e .
if %errorlevel% neq 0 (
    echo ERROR: Failed to install package.
    echo.
    pause
    exit /b 1
)
echo Package installed successfully.
echo.

echo ========================================
echo Setup completed successfully!
echo ========================================
echo.
echo You can now run the application using:
echo   run.bat [input_path] [output_path]
echo.
echo Or activate the virtual environment manually:
echo   .venv\Scripts\activate.bat
echo   docs-to-markdown [input_path] [output_path]
echo.

pause
