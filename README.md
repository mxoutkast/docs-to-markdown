# docs-to-markdown

Batch convert `.docx` (and optionally `.doc`) files to Markdown.

## What it does

- Scans an input folder for `.docx` and `.doc` files
- Converts each to a `.md` file
- Retains base filenames (and mirrors subfolders into the output folder)
- Extracts embedded images into a sibling `<name>_files/` folder

## Requirements

- Python 3.9+
- For `.docx`: no extra system dependencies
- For legacy `.doc`: **LibreOffice** is recommended (used to convert `.doc` → `.docx` headlessly)

## Install

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
