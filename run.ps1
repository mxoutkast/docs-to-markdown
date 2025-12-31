# docs-to-markdown Launcher Script
# PowerShell version for launching the application

$ErrorActionPreference = "Stop"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "docs-to-markdown Launcher" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check if virtual environment exists
if (-not (Test-Path ".venv")) {
    Write-Host "ERROR: Virtual environment not found." -ForegroundColor Red
    Write-Host "Please run setup.ps1 first to create the virtual environment." -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

# Activate virtual environment
$activateScript = Join-Path $PSScriptRoot ".venv\Scripts\Activate.ps1"
if (Test-Path $activateScript) {
    try {
        & $activateScript
    }
    catch {
        Write-Host "ERROR: Failed to activate virtual environment." -ForegroundColor Red
        Write-Host ""
        Read-Host "Press Enter to exit"
        exit 1
    }
}
else {
    Write-Host "ERROR: Failed to find activation script." -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

# Run docs-to-markdown with all arguments
$argsString = $args -join " "
Write-Host "Running: docs-to-markdown $argsString" -ForegroundColor Yellow
Write-Host ""

& docs-to-markdown @args

# Check exit code
if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "Application exited with error code $LASTEXITCODE." -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit $LASTEXITCODE
}

Write-Host ""
Read-Host "Press Enter to exit"
