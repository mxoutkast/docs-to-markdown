# docs-to-markdown

Batch convert `.docx` (and optionally `.doc`) files to Markdown.

## What it does

- Scans an input folder for `.docx` and `.doc` files
- Converts each to a `.md` file
- Retains base filenames (and mirrors subfolders into the output folder)
- Extracts embedded images into a sibling `<name>_files/` folder

## Quick Start (1-Click Setup)

**Windows Users:**

1. Double-click `setup.bat` to automatically create the virtual environment and install all dependencies
2. Double-click `run.bat` to launch the application

**PowerShell Users:**

1. Run `.\setup.ps1` to automatically create the virtual environment and install all dependencies
2. Run `.\run.ps1` to launch the application

**Example usage:**
```powershell
# Using batch files
run.bat "C:\path\to\input" "C:\path\to\output"

# Using PowerShell
.\run.ps1 "C:\path\to\input" "C:\path\to\output"
```

The setup scripts will:
- Check that Python 3.9+ is installed
- Create a virtual environment (`.venv`)
- Install all required dependencies
- Install the `docs-to-markdown` package in editable mode

## Requirements

- Python 3.9+
- For `.docx`: no extra system dependencies
- For legacy `.doc`: **LibreOffice** is recommended (used to convert `.doc` → `.docx` headlessly)

## Install

### Option 1: 1-Click Setup (Recommended)

See **Quick Start** above for automatic setup using `setup.bat` or `setup.ps1`.

### Option 2: Manual Setup

If you prefer manual setup or are on a non-Windows platform:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt

# Install the package so you can run it from any folder
python -m pip install -e .
```

(Optional, for tests)

```powershell
python -m pip install -r requirements-dev.txt
```

## Usage

### Using 1-Click Launcher Scripts (Recommended)

After running the setup script, use the launcher scripts to run the application:

**Windows Batch:**
```powershell
run.bat "C:\path\to\input" "C:\path\to\output"
```

**PowerShell:**
```powershell
.\run.ps1 "C:\path\to\input" "C:\path\to\output"
```

If `OUTPUT_DIR` is omitted, output defaults to a `markdown` subfolder under the input directory:

```powershell
run.bat "C:\path\to\input"
```

Include subfolders:

```powershell
run.bat "C:\path\to\input" -r
```

### Using Command Line Directly

Convert everything under `INPUT_DIR` into `OUTPUT_DIR`:

```powershell
docs-to-markdown "C:\path\to\input" "C:\path\to\output"
```

If `OUTPUT_DIR` is omitted, output defaults to a `markdown` subfolder under the input directory:

```powershell
docs-to-markdown "C:\path\to\input"
```

Include subfolders:

```powershell
docs-to-markdown "C:\path\to\input" -r
```

You can also run it without installing (from the repo root only):

```powershell
python -m docs_to_markdown "C:\path\to\input" "C:\path\to\output"
```

## Notes

- `.docx` conversion uses `mammoth` (DOCX → HTML) then `markdownify` (HTML → Markdown).
- Embedded images are written next to the `.md` as `<md_stem>_files/imageN.<ext>` and linked relatively.
- `.doc` conversion requires LibreOffice (`soffice`) on PATH; otherwise `.doc` files will be skipped with an error message.

## Troubleshooting

### Setup Scripts

**Python not found:**
- Ensure Python 3.9+ is installed from https://www.python.org/
- Make sure to check "Add Python to PATH" during installation
- Restart your terminal/command prompt after installation

**Virtual environment already exists:**
- The setup script will ask if you want to recreate it
- Choose "y" to recreate, or "N" to skip and use the existing one

**Permission errors:**
- Run Command Prompt or PowerShell as Administrator
- Ensure you have write permissions in the project directory

**Network errors during pip install:**
- Check your internet connection
- If behind a corporate firewall, you may need to configure pip with a proxy

### Launcher Scripts

**"Virtual environment not found" error:**
- Run the setup script (`setup.bat` or `setup.ps1`) first
- Ensure the `.venv` directory exists in the project root

**Activation fails:**
- Make sure the virtual environment was created successfully
- Try running the setup script again

**Application errors:**
- Check that your input path is correct
- Ensure you have read permissions for the input directory
- Ensure you have write permissions for the output directory

### PowerShell Execution Policy

If you get an error running `setup.ps1` or `run.ps1` about execution policy, you can bypass it:

```powershell
powershell -ExecutionPolicy Bypass -File setup.ps1
powershell -ExecutionPolicy Bypass -File run.ps1
```

Or change your execution policy (requires administrator privileges):

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```
