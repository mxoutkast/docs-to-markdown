# docs-to-markdown Setup Script
# PowerShell version for automated environment setup

$ErrorActionPreference = "Stop"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "docs-to-markdown Setup Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check if Python is installed
try {
    $pythonVersion = python --version 2>&1
    Write-Host "Checking Python version..." -ForegroundColor Yellow
    Write-Host "Found $pythonVersion" -ForegroundColor Green
    Write-Host ""
}
catch {
    Write-Host "ERROR: Python is not installed or not in PATH." -ForegroundColor Red
    Write-Host "Please install Python 3.9 or higher from https://www.python.org/" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

# Check if .venv already exists
$skipVenvCreation = $false

if (Test-Path ".venv") {
    Write-Host "Virtual environment already exists." -ForegroundColor Yellow
    $recreate = Read-Host "Do you want to recreate it? (y/N)"
    
    if ($recreate -eq "y" -or $recreate -eq "Y") {
        Write-Host "Removing existing virtual environment..." -ForegroundColor Yellow
        try {
            Remove-Item -Path ".venv" -Recurse -Force
        }
        catch {
            Write-Host "ERROR: Failed to remove existing .venv directory." -ForegroundColor Red
            Write-Host "Please close any applications using the virtual environment and try again." -ForegroundColor Yellow
            Write-Host ""
            Read-Host "Press Enter to exit"
            exit 1
        }
    }
    else {
        Write-Host "Skipping virtual environment creation." -ForegroundColor Yellow
        Write-Host ""
        $skipVenvCreation = $true
    }
}

# Create virtual environment (if not skipped)
if (-not $skipVenvCreation) {
    Write-Host "Creating virtual environment..." -ForegroundColor Yellow
    try {
        python -m venv .venv
    }
    catch {
        Write-Host "ERROR: Failed to create virtual environment." -ForegroundColor Red
        Write-Host "Please check that you have Python 3.9+ installed and have sufficient permissions." -ForegroundColor Yellow
        Write-Host ""
        Read-Host "Press Enter to exit"
        exit 1
    }
    Write-Host "Virtual environment created successfully." -ForegroundColor Green
    Write-Host ""
}

# Activate virtual environment and install dependencies
Write-Host "Installing dependencies..." -ForegroundColor Yellow

# Activate the virtual environment
$activateScript = Join-Path $PSScriptRoot ".venv\Scripts\Activate.ps1"
if (Test-Path $activateScript) {
    & $activateScript
}
else {
    Write-Host "ERROR: Failed to find activation script." -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

# Upgrade pip
try {
    python -m pip install --upgrade pip
}
catch {
    Write-Host "ERROR: Failed to upgrade pip." -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

# Install dependencies from requirements.txt
if (Test-Path "requirements.txt") {
    try {
        python -m pip install -r requirements.txt
    }
    catch {
        Write-Host "ERROR: Failed to install dependencies from requirements.txt." -ForegroundColor Red
        Write-Host "Please check your internet connection and try again." -ForegroundColor Yellow
        Write-Host ""
        Read-Host "Press Enter to exit"
        exit 1
    }
    Write-Host "Dependencies installed successfully." -ForegroundColor Green
    Write-Host ""
}
else {
    Write-Host "WARNING: requirements.txt not found. Skipping dependency installation." -ForegroundColor Yellow
    Write-Host ""
}

# Install package in editable mode
Write-Host "Installing docs-to-markdown package..." -ForegroundColor Yellow
try {
    python -m pip install -e .
}
catch {
    Write-Host "ERROR: Failed to install package." -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}
Write-Host "Package installed successfully." -ForegroundColor Green
Write-Host ""

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Setup completed successfully!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "You can now run the application using:" -ForegroundColor White
Write-Host "  .\run.ps1 [input_path] [output_path]" -ForegroundColor Cyan
Write-Host ""
Write-Host "Or activate the virtual environment manually:" -ForegroundColor White
Write-Host "  .\.venv\Scripts\Activate.ps1" -ForegroundColor Cyan
Write-Host "  docs-to-markdown [input_path] [output_path]" -ForegroundColor Cyan
Write-Host ""

Read-Host "Press Enter to exit"
