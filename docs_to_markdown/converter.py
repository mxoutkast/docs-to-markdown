from __future__ import annotations

import shutil
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import BinaryIO

import mammoth
import mammoth.images
from bs4 import BeautifulSoup
from markdownify import MarkdownConverter


@dataclass(frozen=True)
class ConversionReport:
    converted: int
    skipped: int
    failed: int
    failures: tuple[str, ...]


def _iter_input_files(input_dir: Path, *, recursive: bool, include_doc: bool) -> list[Path]:
    if recursive:
        candidates = list(input_dir.rglob("*"))
    else:
        candidates = list(input_dir.glob("*"))

    exts = {".docx"}
    if include_doc:
        exts.add(".doc")

    files: list[Path] = []
    for p in candidates:
        if not p.is_file():
            continue
        if p.suffix.lower() in exts:
            files.append(p)

    files.sort(key=lambda x: str(x).lower())
    return files


def _output_md_path(input_file: Path, *, input_dir: Path, output_dir: Path) -> Path:
    rel = input_file.relative_to(input_dir)
    # Mirror subfolders; keep base filename; change extension to .md
    return (output_dir / rel).with_suffix(".md")


def _ensure_parent_dir(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def _assets_dir_for_md(md_path: Path) -> Path:
    return md_path.parent / f"{md_path.stem}_files"


def _extension_from_content_type(content_type: str) -> str:
    content_type = (content_type or "").lower().strip()
    mapping = {
        "image/png": ".png",
        "image/jpeg": ".jpg",
        "image/jpg": ".jpg",
        "image/gif": ".gif",
        "image/webp": ".webp",
        "image/bmp": ".bmp",
        "image/tiff": ".tiff",
        "image/svg+xml": ".svg",
    }
    if content_type in mapping:
        return mapping[content_type]

    if "/" in content_type:
        subtype = content_type.split("/", 1)[1]
        subtype = subtype.split(";", 1)[0]
        subtype = subtype.split("+", 1)[0]
        subtype = subtype.strip()
        if subtype:
            return "." + subtype

    return ".bin"


def _normalize_html(html: str) -> str:
    # Mammoth emits HTML fragments. Wrap to make soup parsing consistent.
    soup = BeautifulSoup(f"<div>{html}</div>", "lxml")
    wrapper = soup.find("div")
    return "".join(str(child) for child in (wrapper.contents if wrapper else soup.contents))


class _MdConverter(MarkdownConverter):
    # Keep markdownify defaults, but avoid overly aggressive escaping.
    def __init__(self) -> None:
        super().__init__(
            heading_style="ATX",
            bullets="-",
            strong_em_symbol="*",
            code_language="",
        )


def _docx_to_markdown(docx_path: Path, *, output_md_path: Path) -> str:
    images_written = 0
    image_index = 0
    assets_dir = _assets_dir_for_md(output_md_path)

    def _convert_image(image: object) -> dict[str, str]:
        nonlocal image_index, images_written
        image_index += 1

        content_type = getattr(image, "content_type", "")
        ext = _extension_from_content_type(str(content_type))
        filename = f"image{image_index}{ext}"
        file_path = assets_dir / filename

        assets_dir.mkdir(parents=True, exist_ok=True)

        opener = getattr(image, "open")
        with opener() as image_bytes:  # type: ignore[call-arg]
            if isinstance(image_bytes, (bytes, bytearray)):
                data = bytes(image_bytes)
            else:
                data = image_bytes.read()  # type: ignore[union-attr]
        file_path.write_bytes(data)
        images_written += 1

        # Use POSIX separators for Markdown links.
        rel = Path(assets_dir.name) / filename
        return {"src": rel.as_posix()}

    with docx_path.open("rb") as f:
        result = mammoth.convert_to_html(
            f,
            convert_image=mammoth.images.img_element(_convert_image),
        )

    html = _normalize_html(result.value)
    md = _MdConverter().convert(html)

    # mammoth/markdownify can produce excessive blank lines; keep it tidy.
    md = md.replace("\r\n", "\n")
    while "\n\n\n" in md:
        md = md.replace("\n\n\n", "\n\n")

    # If no images were written, avoid leaving an empty assets directory.
    if images_written == 0:
        try:
            if assets_dir.exists():
                assets_dir.rmdir()
        except Exception:
            pass

    return md.strip() + "\n"


def _find_soffice() -> str | None:
    return shutil.which("soffice") or shutil.which("soffice.exe")


def _doc_to_docx_via_libreoffice(doc_path: Path, *, work_dir: Path) -> Path:
    soffice = _find_soffice()
    if not soffice:
        raise RuntimeError("LibreOffice 'soffice' not found on PATH; cannot convert .doc")

    # LibreOffice writes output into --outdir with same base filename.
    outdir = work_dir
    outdir.mkdir(parents=True, exist_ok=True)

    cmd = [
        soffice,
        "--headless",
        "--nologo",
        "--nolockcheck",
        "--nodefault",
        "--norestore",
        "--convert-to",
        "docx",
        "--outdir",
        str(outdir),
        str(doc_path),
    ]

    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        stderr = (proc.stderr or proc.stdout or "").strip()
        raise RuntimeError(f"LibreOffice conversion failed: {stderr}")

    produced = outdir / (doc_path.stem + ".docx")
    if not produced.exists():
        raise RuntimeError("LibreOffice did not produce expected .docx output")

    return produced


def convert_folder(
    *,
    input_dir: Path,
    output_dir: Path,
    recursive: bool = False,
    include_doc: bool = False,
    overwrite: bool = False,
) -> ConversionReport:
    input_dir = input_dir.resolve()
    output_dir = output_dir.resolve()

    if not input_dir.exists() or not input_dir.is_dir():
        raise ValueError(f"input_dir does not exist or is not a directory: {input_dir}")

    files = _iter_input_files(input_dir, recursive=recursive, include_doc=include_doc)

    converted = 0
    skipped = 0
    failed = 0
    failures: list[str] = []

    temp_dir = output_dir / ".__tmp_doc_conversion__"

    for src in files:
        dest = _output_md_path(src, input_dir=input_dir, output_dir=output_dir)

        if dest.exists() and not overwrite:
            skipped += 1
            continue

        try:
            _ensure_parent_dir(dest)

            if src.suffix.lower() == ".docx":
                md = _docx_to_markdown(src, output_md_path=dest)
            else:
                # .doc path
                docx_path = _doc_to_docx_via_libreoffice(src, work_dir=temp_dir)
                md = _docx_to_markdown(docx_path, output_md_path=dest)

            dest.write_text(md, encoding="utf-8")
            converted += 1
        except Exception as e:
            failed += 1
            failures.append(f"{src}: {e}")

    # Best-effort cleanup
    try:
        if temp_dir.exists():
            shutil.rmtree(temp_dir)
    except Exception:
        pass

    return ConversionReport(
        converted=converted,
        skipped=skipped,
        failed=failed,
        failures=tuple(failures),
    )
