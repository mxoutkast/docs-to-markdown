from __future__ import annotations

import shutil
import subprocess
import sys
import threading
from dataclasses import dataclass
from pathlib import Path
from typing import BinaryIO, Callable

import customtkinter as ctk
from tkinter import filedialog, messagebox
import tkinterdnd2 as tkdnd

import mammoth
import mammoth.images
from bs4 import BeautifulSoup
from markdownify import MarkdownConverter
import markdown

from docs_to_markdown.converter import (
    _doc_to_docx_via_libreoffice,
    _docx_to_markdown,
    _ensure_parent_dir,
    _extension_from_content_type,
    _iter_input_files,
    _normalize_html,
    _output_md_path,
    ConversionReport,
)


def convert_folder_with_progress(
    *,
    input_dir: Path,
    output_dir: Path,
    recursive: bool = False,
    include_doc: bool = False,
    overwrite: bool = False,
    progress_callback: Callable[[int, int, Path], None] | None = None,
    stop_event: threading.Event | None = None,
) -> ConversionReport:
    """
    Convert files with progress tracking.

    Args:
        input_dir: Input directory containing .doc/.docx files.
        output_dir: Output directory for .md files.
        recursive: Whether to search subfolders.
        include_doc: Whether to include .doc files.
        overwrite: Whether to overwrite existing .md files.
        progress_callback: Optional callback(current, total, current_file) for progress updates.
        stop_event: Optional threading.Event to signal cancellation.

    Returns:
        ConversionReport with conversion statistics.
    """
    input_dir = input_dir.resolve()
    output_dir = output_dir.resolve()

    if not input_dir.exists() or not input_dir.is_dir():
        raise ValueError(f"input_dir does not exist or is not a directory: {input_dir}")

    files = _iter_input_files(input_dir, recursive=recursive, include_doc=include_doc)
    total_files = len(files)

    converted = 0
    skipped = 0
    failed = 0
    failures: list[str] = []

    temp_dir = output_dir / ".__tmp_doc_conversion__"

    for idx, src in enumerate(files):
        # Check for cancellation
        if stop_event and stop_event.is_set():
            break

        # Update progress
        if progress_callback:
            progress_callback(idx, total_files, src)

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

    # Final progress update
    if progress_callback:
        progress_callback(total_files, total_files, Path(""))

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


class DocsToMarkdownGUI:
    """Main GUI window for docs-to-markdown converter."""

    def __init__(self) -> None:
        """Initialize the GUI application."""
        self._root = tkdnd.Tk()
        self._root.title("Docs to Markdown Converter")
        self._root.geometry("900x700")
        self._root.minsize(700, 600)

        # Configure customtkinter theme
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        # Configure grid weights for proper resizing
        self._root.grid_rowconfigure(0, weight=1)
        self._root.grid_columnconfigure(0, weight=1)

        # State variables
        self._input_path: str = ""
        self._output_path: str = ""
        self._last_output_path: Path | None = None
        self._recursive: bool = False
        self._include_doc: bool = False
        self._overwrite: bool = False
        self._conversion_thread: threading.Thread | None = None
        self._stop_conversion = threading.Event()
        self._current_markdown_text: str = ""

        # Build UI
        self._setup_layout()
        self._setup_drag_drop()

    def _render_markdown(self, md_text: str) -> str:
        """
        Convert markdown text to formatted text for display in preview.

        Args:
            md_text: Raw markdown text.

        Returns:
            Formatted text with proper formatting for display.
        """
        if not md_text:
            return ""

        # Convert markdown to HTML with extensions for better formatting
        html = markdown.markdown(
            md_text,
            extensions=[
                "extra",  # Tables, fenced code blocks, etc.
                "nl2br",  # New line to <br>
                "sane_lists",  # Better list handling
            ],
        )

        # Parse HTML and convert to formatted text
        soup = BeautifulSoup(html, "lxml")

        # Build formatted text with tags for CTkTextbox
        formatted_lines = []

        # Helper function to recursively process elements
        def process_element(element, depth=0):
            """Recursively process HTML elements to build formatted text."""
            lines = []

            # Handle different element types
            if element.name in ["h1", "h2", "h3", "h4", "h5", "h6"]:
                # Headers
                level = int(element.name[1])
                prefix = "#" * level + " "
                text = element.get_text(strip=True)
                if text:
                    lines.append(f"{prefix}{text}")
                    lines.append("")  # Add blank line after header
            elif element.name == "p":
                # Paragraph - process children inline
                text_parts = []
                for child in element.children:
                    if hasattr(child, 'name'):
                        if child.name in ["strong", "b"]:
                            text_parts.append(f"**{child.get_text(strip=True)}**")
                        elif child.name in ["em", "i"]:
                            text_parts.append(f"*{child.get_text(strip=True)}*")
                        elif child.name == "code":
                            text_parts.append(f"`{child.get_text(strip=True)}`")
                        elif child.name == "a":
                            href = child.get("href", "")
                            text = child.get_text(strip=True)
                            if href:
                                text_parts.append(f"[{text}]({href})")
                            else:
                                text_parts.append(text)
                        else:
                            text_parts.append(child.get_text(strip=True))
                    else:
                        # Text node
                        text = str(child).strip()
                        if text:
                            text_parts.append(text)

                paragraph_text = " ".join(text_parts)
                if paragraph_text:
                    lines.append(paragraph_text)
                    lines.append("")  # Add blank line after paragraph
            elif element.name == "ul" or element.name == "ol":
                # Lists
                for li in element.find_all("li", recursive=False):
                    text = li.get_text(strip=True)
                    if text:
                        lines.append(f"- {text}")
            elif element.name == "pre":
                # Code block
                code_text = element.get_text()
                if code_text.strip():
                    lines.append("```")
                    for line in code_text.split("\n"):
                        lines.append(f"    {line}")
                    lines.append("```")
            elif element.name == "blockquote":
                # Blockquote
                text = element.get_text(strip=True)
                if text:
                    lines.append(f"> {text}")
                    lines.append("")
            elif element.name == "hr":
                # Horizontal rule
                lines.append("---")
                lines.append("")
            elif element.name == "table":
                # Table (basic support)
                rows = element.find_all("tr")
                if rows:
                    for row in rows:
                        cells = row.find_all(["th", "td"])
                        if cells:
                            row_text = " | ".join(cell.get_text(strip=True) for cell in cells)
                            lines.append(f"| {row_text} |")
                    lines.append("")

            return lines

        # Find the body element (BeautifulSoup wraps HTML in html/body tags)
        body = soup.find("body")
        if body:
            # Process all top-level elements in body
            for element in body.children:
                if hasattr(element, 'name'):
                    formatted_lines.extend(process_element(element))
                else:
                    # Text node at top level
                    text = str(element).strip()
                    if text:
                        formatted_lines.append(text)
        else:
            # Fallback: process soup children directly
            for element in soup.children:
                if hasattr(element, 'name'):
                    formatted_lines.extend(process_element(element))
                else:
                    text = str(element).strip()
                    if text:
                        formatted_lines.append(text)

        # Join lines with newlines
        formatted_text = "\n".join(formatted_lines)

        # Clean up excessive blank lines
       
... (truncated)