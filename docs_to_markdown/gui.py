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
        while "\n\n\n" in formatted_text:
            formatted_text = formatted_text.replace("\n\n\n", "\n\n")

        return formatted_text.strip()

    def _show_preview(self, md_path: Path) -> None:
        """
        Display markdown file content in preview panel with proper formatting.

        Args:
            md_path: Path to markdown file to preview.
        """
        try:
            if not md_path.exists() or not md_path.is_file():
                messagebox.showerror(
                    "Error",
                    f"File not found: {md_path}",
                )
                return

            # Read markdown content
            md_text = md_path.read_text(encoding="utf-8")

            # Render markdown with proper formatting
            formatted_text = self._render_markdown(md_text)

            # Update preview textbox
            self._preview_text.configure(state="normal")
            self._preview_text.delete("1.0", "end")
            self._preview_text.insert("1.0", formatted_text)
            self._preview_text.configure(state="disabled")

        except Exception as e:
            messagebox.showerror(
                "Error",
                f"Failed to preview file: {e}",
            )

    def _setup_layout(self) -> None:
        """Create and configure the main window layout."""
        # Main container with scrollable frame
        self._main_frame = ctk.CTkScrollableFrame(
            self._root,
            label_text="Docs to Markdown Converter",
        )
        self._main_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self._main_frame.grid_rowconfigure(0, weight=0)  # Input section
        self._main_frame.grid_rowconfigure(1, weight=0)  # Output section
        self._main_frame.grid_rowconfigure(2, weight=0)  # Options section
        self._main_frame.grid_rowconfigure(3, weight=0)  # Convert section
        self._main_frame.grid_rowconfigure(4, weight=0)  # Progress section
        self._main_frame.grid_rowconfigure(5, weight=0)  # Results section
        self._main_frame.grid_rowconfigure(6, weight=1)  # Preview section (expands)
        self._main_frame.grid_columnconfigure(0, weight=1)

        # Input path section
        self._create_input_section()

        # Output path section
        self._create_output_section()

        # Options section
        self._create_options_section()

        # Convert button section
        self._create_convert_section()

        # Progress section
        self._create_progress_section()

        # Results section
        self._create_results_section()

        # Preview section
        self._create_preview_section()

    def _create_input_section(self) -> None:
        """Create input path selection section."""
        input_frame = ctk.CTkFrame(self._main_frame)
        input_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))

        input_label = ctk.CTkLabel(
            input_frame,
            text="Input File/Folder:",
            font=("Helvetica", 12, "bold"),
        )
        input_label.pack(anchor="w", padx=10, pady=(10, 5))

        input_container = ctk.CTkFrame(input_frame)
        input_container.pack(fill="x", padx=10, pady=(0, 10))

        self._input_entry = ctk.CTkEntry(
            input_container,
            placeholder_text="Select .doc/.docx file or folder containing files",
        )
        self._input_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))

        input_browse_btn = ctk.CTkButton(
            input_container,
            text="Browse",
            width=80,
            command=self._browse_input,
        )
        input_browse_btn.pack(side="right")

    def _create_output_section(self) -> None:
        """Create output path selection section."""
        output_frame = ctk.CTkFrame(self._main_frame)
        output_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))

        output_label = ctk.CTkLabel(
            output_frame,
            text="Output Folder:",
            font=("Helvetica", 12, "bold"),
        )
        output_label.pack(anchor="w", padx=10, pady=(10, 5))

        output_container = ctk.CTkFrame(output_frame)
        output_container.pack(fill="x", padx=10, pady=(0, 10))

        self._output_entry = ctk.CTkEntry(
            output_container,
            placeholder_text="Select output folder (default: <input>/markdown)",
        )
        self._output_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))

        output_browse_btn = ctk.CTkButton(
            output_container,
            text="Browse",
            width=80,
            command=self._browse_output,
        )
        output_browse_btn.pack(side="right")

    def _create_options_section(self) -> None:
        """Create conversion options checkboxes."""
        options_frame = ctk.CTkFrame(self._main_frame)
        options_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))

        options_label = ctk.CTkLabel(
            options_frame,
            text="Options:",
            font=("Helvetica", 12, "bold"),
        )
        options_label.pack(anchor="w", padx=10, pady=(10, 5))

        checkboxes_frame = ctk.CTkFrame(options_frame)
        checkboxes_frame.pack(fill="x", padx=10, pady=(0, 10))

        self._recursive_var = ctk.BooleanVar(value=False)
        recursive_cb = ctk.CTkCheckBox(
            checkboxes_frame,
            text="Recursive (include subfolders)",
            variable=self._recursive_var,
            command=self._update_options,
        )
        recursive_cb.pack(anchor="w", padx=10, pady=5)

        self._include_doc_var = ctk.BooleanVar(value=False)
        include_doc_cb = ctk.CTkCheckBox(
            checkboxes_frame,
            text="Include .doc files (requires LibreOffice)",
            variable=self._include_doc_var,
            command=self._update_options,
        )
        include_doc_cb.pack(anchor="w", padx=10, pady=5)

        self._overwrite_var = ctk.BooleanVar(value=False)
        overwrite_cb = ctk.CTkCheckBox(
            checkboxes_frame,
            text="Overwrite existing .md files",
            variable=self._overwrite_var,
            command=self._update_options,
        )
        overwrite_cb.pack(anchor="w", padx=10, pady=5)

    def _create_convert_section(self) -> None:
        """Create convert button section."""
        button_frame = ctk.CTkFrame(self._main_frame)
        button_frame.grid(row=3, column=0, sticky="ew", pady=(0, 10))

        self._convert_btn = ctk.CTkButton(
            button_frame,
            text="Convert",
            font=("Helvetica", 14, "bold"),
            height=40,
            command=self._start_conversion,
        )
        self._convert_btn.pack(fill="x", padx=10, pady=10)

        self._cancel_btn = ctk.CTkButton(
            button_frame,
            text="Cancel",
            font=("Helvetica", 14, "bold"),
            height=40,
            command=self._cancel_conversion,
            state="disabled",
        )
        self._cancel_btn.pack(fill="x", padx=10, pady=(0, 10))

    def _create_progress_section(self) -> None:
        """Create progress indicator section."""
        progress_frame = ctk.CTkFrame(self._main_frame)
        progress_frame.grid(row=4, column=0, sticky="ew", pady=(0, 10))

        progress_label = ctk.CTkLabel(
            progress_frame,
            text="Progress:",
            font=("Helvetica", 12, "bold"),
        )
        progress_label.pack(anchor="w", padx=10, pady=(10, 5))

        self._progress_bar = ctk.CTkProgressBar(progress_frame)
        self._progress_bar.pack(fill="x", padx=10, pady=(0, 5))
        self._progress_bar.set(0)

        self._status_label = ctk.CTkLabel(
            progress_frame,
            text="Ready",
            text_color="gray",
        )
        self._status_label.pack(anchor="w", padx=10, pady=(0, 10))

    def _create_results_section(self) -> None:
        """Create results display section."""
        results_frame = ctk.CTkFrame(self._main_frame)
        results_frame.grid(row=5, column=0, sticky="ew", pady=(0, 10))

        results_label = ctk.CTkLabel(
            results_frame,
            text="Results:",
            font=("Helvetica", 12, "bold"),
        )
        results_label.pack(anchor="w", padx=10, pady=(10, 5))

        self._results_text = ctk.CTkTextbox(
            results_frame,
            height=100,
        )
        self._results_text.pack(fill="x", padx=10, pady=(0, 10))
        self._results_text.insert("1.0", "Conversion results will appear here...")
        self._results_text.configure(state="disabled")

    def _create_preview_section(self) -> None:
        """Create markdown preview section with file browser."""
        preview_frame = ctk.CTkFrame(self._main_frame)
        preview_frame.grid(row=6, column=0, sticky="nsew", pady=(0, 10))
        preview_frame.grid_rowconfigure(1, weight=1)  # Make preview textbox expand
        preview_frame.grid_columnconfigure(0, weight=1)

        preview_label = ctk.CTkLabel(
            preview_frame,
            text="Markdown Preview:",
            font=("Helvetica", 12, "bold"),
        )
        preview_label.grid(row=0, column=0, sticky="w", padx=10, pady=(10, 5))

        # File browser frame
        file_browser_frame = ctk.CTkFrame(preview_frame)
        file_browser_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 10))

        file_browser_label = ctk.CTkLabel(
            file_browser_frame,
            text="Converted Files:",
            font=("Helvetica", 10, "bold"),
        )
        file_browser_label.pack(anchor="w", padx=5, pady=(5, 0))

        # Create scrollable listbox for converted files
        self._files_listbox = ctk.CTkScrollableFrame(
            file_browser_frame,
            height=100,
        )
        self._files_listbox.pack(fill="x", padx=5, pady=5)

        # Preview textbox
        self._preview_text = ctk.CTkTextbox(
            preview_frame,
            height=200,
        )
        self._preview_text.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self._preview_text.insert("1.0", "Preview converted markdown here...")
        self._preview_text.configure(state="disabled")

    def _update_files_list(self, output_path: Path) -> None:
        """
        Update the files listbox with converted markdown files.

        Args:
            output_path: Path to output directory containing converted files.
        """
        # Clear existing file buttons
        for widget in self._files_listbox.winfo_children():
            widget.destroy()

        if not output_path.exists() or not output_path.is_dir():
            return

        # Find all .md files in output directory
        md_files = list(output_path.rglob("*.md"))
        md_files.sort(key=lambda x: str(x).lower())

        # Add clickable buttons for each file
        for md_file in md_files:
            # Calculate relative path for display
            try:
                rel_path = md_file.relative_to(output_path)
            except ValueError:
                rel_path = md_file.name

            file_btn = ctk.CTkButton(
                self._files_listbox,
                text=str(rel_path),
                command=lambda f=md_file: self._show_preview(f),
            )
            file_btn.pack(fill="x", pady=2)

    def _setup_drag_drop(self) -> None:
        """Configure drag-and-drop functionality."""
        self._root.drop_target_register(tkdnd.DND_FILES)
        self._root.dnd_bind("<<Drop>>", self._on_drop)

    def _browse_input(self) -> None:
        """Open file dialog to select input file or folder."""
        try:
            # Show dialog to choose between file or folder selection
            choice = messagebox.askyesno(
                "Select Input",
                "Do you want to select a single file?\n\n"
                "Yes = Select a single .doc/.docx file\n"
                "No = Select a folder containing files",
            )

            if choice:
                # File selection
                file_path = filedialog.askopenfilename(
                    title="Select Input File",
                    filetypes=[
                        ("Word Documents", "*.docx *.doc"),
                        ("Word Document (.docx)", "*.docx"),
                        ("Word Document (.doc)", "*.doc"),
                        ("All Files", "*.*"),
                    ],
                )
                if file_path:
                    input_path = Path(file_path).resolve()

                    # Validate file exists
                    if not input_path.exists():
                        messagebox.showerror(
                            "Error",
                            f"File does not exist: {file_path}",
                        )
                        return

                    # Validate it's a file
                    if not input_path.is_file():
                        messagebox.showerror(
                            "Error",
                            f"Path is not a file: {file_path}",
                        )
                        return

                    # Validate file type
                    ext = input_path.suffix.lower()
                    if ext not in {".doc", ".docx"}:
                        messagebox.showerror(
                            "Invalid File Type",
                            f"Selected file is not a Word document (.doc/.docx).\n\n"
                            f"File: {input_path.name}\n"
                            f"Type: {ext or 'unknown'}\n\n"
                            f"Please select a valid Word document (.doc or .docx).",
                        )
                        # Don't allow proceeding with invalid file type
                        return

                    # If it's a .doc file, check LibreOffice availability
                    if ext == ".doc":
                        soffice = shutil.which("soffice") or shutil.which("soffice.exe")
                        if not soffice:
                            messagebox.showwarning(
                                "LibreOffice Not Found",
                                f"Selected file is a .doc file, but LibreOffice was not found.\n\n"
                                f"File: {input_path.name}\n\n"
                                f"LibreOffice is required to convert .doc files.\n\n"
                                f"Please:\n"
                                f"- Install LibreOffice: https://www.libreoffice.org/download/\n"
                                f"- Or select a .docx file instead\n\n"
                                f"You can still proceed, but the conversion will fail.",
                            )

                    self._input_entry.delete(0, "end")
                    self._input_entry.insert(0, str(input_path))
                    self._input_path = str(input_path)
            else:
                # Folder selection
                folder = filedialog.askdirectory(title="Select Input Folder")
                if folder:
                    input_path = Path(folder).resolve()

                    # Validate folder exists
                    if not input_path.exists():
                        messagebox.showerror(
                            "Error",
                            f"Folder does not exist: {folder}",
                        )
                        return

                    # Validate it's a directory
                    if not input_path.is_dir():
                        messagebox.showerror(
                            "Error",
                            f"Path is not a directory: {folder}",
                        )
                        return

                    # Check if folder contains valid files
                    has_docx = any(input_path.glob("*.docx")) or any(input_path.glob("*.DOCX"))
                    has_doc = any(input_path.glob("*.doc")) or any(input_path.glob("*.DOC"))

                    if not has_docx and not has_doc:
                        messagebox.showerror(
                            "No Valid Files Found",
                            f"Selected folder does not contain any Word documents (.doc/.docx).\n\n"
                            f"Folder: {input_path}\n\n"
                            f"Please select a folder containing Word documents.",
                        )
                        return

                    # If folder contains .doc files, warn about LibreOffice
                    if has_doc:
                        soffice = shutil.which("soffice") or shutil.which("soffice.exe")
                        if not soffice:
                            messagebox.showwarning(
                                "LibreOffice Not Found",
                                f"Selected folder contains .doc files, but LibreOffice was not found.\n\n"
                                f"Folder: {input_path}\n\n"
                                f"LibreOffice is required to convert .doc files.\n\n"
                                f"Please:\n"
                                f"- Install LibreOffice: https://www.libreoffice.org/download/\n"
                                f"- Or enable the 'Include .doc files' option only if LibreOffice is installed\n\n"
                                f"Note: Only .docx files will be converted without LibreOffice.",
                            )

                    self._input_entry.delete(0, "end")
                    self._input_entry.insert(0, str(input_path))
                    self._input_path = str(input_path)
        except Exception as e:
            messagebox.showerror(
                "Error",
                f"Failed to select input: {e}",
            )

    def _browse_output(self) -> None:
        """Open file dialog to select output folder."""
        try:
            folder = filedialog.askdirectory(title="Select Output Folder")
            if folder:
                output_path = Path(folder).resolve()

                # If folder doesn't exist, try to create it
                if not output_path.exists():
                    try:
                        output_path.mkdir(parents=True, exist_ok=True)
                    except PermissionError:
                        messagebox.showerror(
                            "Permission Error",
                            f"Cannot create output folder due to insufficient permissions:\n{folder}",
                        )
                        return
                    except Exception as e:
                        messagebox.showerror(
                            "Error",
                            f"Cannot create output folder: {e}",
                        )
                        return
                else:
                    # Folder exists, validate it's a directory
                    if not output_path.is_dir():
                        messagebox.showerror(
                            "Error",
                            f"Output path exists but is not a directory:\n{folder}",
                        )
                        return

                # Check write permissions
                try:
                    test_file = output_path / ".__write_test__"
                    test_file.touch()
                    test_file.unlink()
                except PermissionError:
                    messagebox.showerror(
                        "Permission Error",
                        f"Cannot write to output folder due to insufficient permissions:\n{folder}",
                    )
                    return
                except Exception as e:
                    messagebox.showerror(
                        "Error",
                        f"Cannot write to output folder: {e}",
                    )
                    return

                self._output_entry.delete(0, "end")
                self._output_entry.insert(0, str(output_path))
                self._output_path = str(output_path)
        except Exception as e:
            messagebox.showerror(
                "Error",
                f"Failed to select output folder: {e}",
            )

    def _update_options(self) -> None:
        """Update options from checkbox values."""
        self._recursive = self._recursive_var.get()
        self._include_doc = self._include_doc_var.get()
        self._overwrite = self._overwrite_var.get()

    def _on_drop(self, event: object) -> None:
        """Handle drag-and-drop event."""
        try:
            # Extract the dropped path from the event
            data = getattr(event, "data", "")
            if not data:
                return

            # Clean up the path (remove curly braces and quotes if present)
            dropped_path = data.strip("{}").strip('"').strip("'")

            if not dropped_path:
                return

            input_path = Path(dropped_path).resolve()

            # Validate the path exists
            if not input_path.exists():
                messagebox.showerror(
                    "Error",
                    f"Path does not exist: {dropped_path}",
                )
                return

            # Handle file or folder
            if input_path.is_file():
                # Validate file type
                ext = input_path.suffix.lower()
                if ext not in {".doc", ".docx"}:
                    messagebox.showerror(
                        "Invalid File Type",
                        f"Dragged file is not a Word document (.doc/.docx).\n\n"
                        f"File: {input_path.name}\n"
                        f"Type: {ext or 'unknown'}\n\n"
                        f"Please drag and drop a valid Word document (.doc or .docx).",
                    )
                    # Don't accept invalid file types
                    return

                # If it's a .doc file, check LibreOffice availability
                if ext == ".doc":
                    soffice = shutil.which("soffice") or shutil.which("soffice.exe")
                    if not soffice:
                        messagebox.showwarning(
                            "LibreOffice Not Found",
                            f"Dragged file is a .doc file, but LibreOffice was not found.\n\n"
                            f"File: {input_path.name}\n\n"
                            f"LibreOffice is required to convert .doc files.\n\n"
                            f"Please:\n"
                            f"- Install LibreOffice: https://www.libreoffice.org/download/\n"
                            f"- Or drag and drop a .docx file instead\n\n"
                            f"You can still proceed, but the conversion will fail.",
                        )

                self._input_entry.delete(0, "end")
                self._input_entry.insert(0, str(input_path))
                self._input_path = str(input_path)
            elif input_path.is_dir():
                # Check if folder contains valid files
                has_docx = any(input_path.glob("*.docx")) or any(input_path.glob("*.DOCX"))
                has_doc = any(input_path.glob("*.doc")) or any(input_path.glob("*.DOC"))

                if not has_docx and not has_doc:
                    messagebox.showerror(
                        "No Valid Files Found",
                        f"Dragged folder does not contain any Word documents (.doc/.docx).\n\n"
                        f"Folder: {input_path}\n\n"
                        f"Please drag and drop a folder containing Word documents.",
                    )
                    return

                # If folder contains .doc files, warn about LibreOffice
                if has_doc:
                    soffice = shutil.which("soffice") or shutil.which("soffice.exe")
                    if not soffice:
                        messagebox.showwarning(
                            "LibreOffice Not Found",
                            f"Dragged folder contains .doc files, but LibreOffice was not found.\n\n"
                            f"Folder: {input_path}\n\n"
                            f"LibreOffice is required to convert .doc files.\n\n"
                            f"Please:\n"
                            f"- Install LibreOffice: https://www.libreoffice.org/download/\n"
                            f"- Or enable the 'Include .doc files' option only if LibreOffice is installed\n\n"
                            f"Note: Only .docx files will be converted without LibreOffice.",
                        )

                self._input_entry.delete(0, "end")
                self._input_entry.insert(0, str(input_path))
                self._input_path = str(input_path)
            else:
                messagebox.showerror(
                    "Error",
                    f"Path is neither a file nor a directory: {dropped_path}",
                )
        except Exception as e:
            messagebox.showerror(
                "Error",
                f"Failed to handle dropped item: {e}",
            )

    def _validate_inputs(
        self,
        input_path: Path,
        output_path: Path,
        include_doc: bool,
    ) -> tuple[bool, str]:
        """
        Validate all inputs before starting conversion.

        Args:
            input_path: Path to input file or folder.
            output_path: Path to output folder.
            include_doc: Whether .doc files should be included.

        Returns:
            Tuple of (is_valid, error_message). If is_valid is True, error_message is empty.
        """
        # Validate input path exists
        if not input_path.exists():
            return False, f"Input path does not exist: {input_path}"

        # Validate input path is either file or directory
        if not input_path.is_file() and not input_path.is_dir():
            return False, f"Input path is neither a file nor a directory: {input_path}"

        # If input is a file, validate it's a Word document
        if input_path.is_file():
            ext = input_path.suffix.lower()
            if ext not in {".doc", ".docx"}:
                return False, (
                    f"Input file must be a Word document (.doc or .docx): {input_path.name}\n\n"
                    f"Selected file type: {ext or 'unknown'}\n"
                    f"Supported types: .doc, .docx"
                )

            # If it's a .doc file, validate LibreOffice is available
            if ext == ".doc" or include_doc:
                soffice = shutil.which("soffice") or shutil.which("soffice.exe")
                if not soffice:
                    return False, (
                        f"LibreOffice is required to convert .doc files but was not found on PATH.\n\n"
                        f"Please install LibreOffice and ensure 'soffice' is available, or select a .docx file instead.\n\n"
                        f"Download LibreOffice: https://www.libreoffice.org/download/"
                    )

        # Validate LibreOffice is available if .doc files are included
        if include_doc:
            soffice = shutil.which("soffice") or shutil.which("soffice.exe")
            if not soffice:
                return False, (
                    "LibreOffice is required to convert .doc files but was not found on PATH.\n\n"
                    "Please install LibreOffice and ensure 'soffice' is available, or disable 'Include .doc files' option.\n\n"
                    "Download LibreOffice: https://www.libreoffice.org/download/"
                )

        # Validate output path is a directory (or can be created as one)
        if output_path.exists() and not output_path.is_dir():
            return False, f"Output path exists but is not a directory: {output_path}"

        # Validate output path is not inside input path (to avoid circular references)
        try:
            output_path.relative_to(input_path)
            return False, (
                f"Output folder cannot be inside the input folder.\n\n"
                f"Please choose a different output location."
            )
        except ValueError:
            # output_path is not relative to input_path, which is good
            pass

        # Validate output path can be written to
        try:
            if not output_path.exists():
                output_path.mkdir(parents=True, exist_ok=True)
            # Try to create a test file to verify write permissions
            test_file = output_path / ".__write_test__"
            test_file.touch()
            test_file.unlink()
        except PermissionError:
            return False, (
                f"Cannot write to output folder due to insufficient permissions.\n\n"
                f"Output folder: {output_path}\n\n"
                f"Possible solutions:\n"
                f"- Choose a different output folder\n"
                f"- Run the application with administrator privileges\n"
                f"- Check folder permissions in Windows Explorer"
            )
        except OSError as e:
            if "read-only" in str(e).lower() or "permission denied" in str(e).lower():
                return False, (
                    f"Cannot write to output folder - the location may be read-only.\n\n"
                    f"Output folder: {output_path}\n\n"
                    f"Possible solutions:\n"
                    f"- Choose a different output folder\n"
                    f"- Run the application with administrator privileges\n"
                    f"- Check folder properties to ensure it's not marked as read-only"
                )
            return False, f"Cannot access output folder: {e}"
        except Exception as e:
            return False, f"Cannot access output folder: {e}"

        # If input is a directory, validate it contains valid files
        if input_path.is_dir():
            # Check if directory contains any .docx files
            has_docx = any(input_path.glob("*.docx")) or any(input_path.glob("*.DOCX"))
            has_doc = any(input_path.glob("*.doc")) or any(input_path.glob("*.DOC"))

            if not has_docx and not has_doc:
                return False, (
                    f"Input folder does not contain any Word documents (.doc/.docx).\n\n"
                    f"Input folder: {input_path}\n\n"
                    f"Please select a folder containing Word documents or select a single file."
                )

            # If .doc files exist but include_doc is not enabled, warn user
            if has_doc and not include_doc:
                return False, (
                    f"Input folder contains .doc files but 'Include .doc files' option is not enabled.\n\n"
                    f"Either:\n"
                    f"- Enable the 'Include .doc files' option (requires LibreOffice)\n"
                    f"- Remove .doc files from the folder\n\n"
                    f"Note: Only .docx files will be converted with current settings."
                )

        return True, ""

    def _start_conversion(self) -> None:
        """Start the conversion process."""
        # Validate input path
        input_str = self._input_entry.get().strip()
        if not input_str:
            messagebox.showerror("Error", "Please select an input file or folder.")
            return

        input_path = Path(input_str).resolve()

        # Determine output path
        output_str = self._output_entry.get().strip()
        if output_str:
            output_path = Path(output_str).resolve()
        else:
            # Default: create "markdown" subfolder next to input
            if input_path.is_file():
                output_path = input_path.parent / "markdown"
            else:
                output_path = input_path / "markdown"

        # Capture options BEFORE starting thread (to avoid Tkinter threading issues)
        self._update_options()
        recursive = self._recursive
        include_doc = self._include_doc
        overwrite = self._overwrite

        # Validate all inputs
        is_valid, error_msg = self._validate_inputs(
            input_path,
            output_path,
            include_doc,
        )
        if not is_valid:
            messagebox.showerror("Input Validation Error", error_msg)
            return

        # Create output directory if needed
        try:
            output_path.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Error", f"Cannot create output folder: {e}")
            return

        # Update UI state
        self._convert_btn.configure(state="disabled")
        self._cancel_btn.configure(state="normal")
        self._progress_bar.set(0)
        self._status_label.configure(text="Starting conversion...", text_color="blue")
        self._results_text.configure(state="normal")
        self._results_text.delete("1.0", "end")
        self._results_text.insert("1.0", "Conversion in progress...\n")
        self._results_text.configure(state="disabled")
        self._stop_conversion.clear()

        # Start conversion in background thread
        self._conversion_thread = threading.Thread(
            target=self._run_conversion,
            args=(input_path, output_path, recursive, include_doc, overwrite),
            daemon=True,
        )
        self._conversion_thread.start()

    def _run_conversion(
        self,
        input_path: Path,
        output_path: Path,
        recursive: bool,
        include_doc: bool,
        overwrite: bool,
    ) -> None:
        """Run conversion in background thread and update UI."""
        try:
            # Save output path for file browser
            self._last_output_path = output_path

            if input_path.is_file():
                # Single file conversion - create temp folder structure
                temp_dir = output_path / ".__temp_single_file__"
                temp_dir.mkdir(parents=True, exist_ok=True)

                # Copy file to temp dir to use convert_folder
                temp_file = temp_dir / input_path.name
                try:
                    shutil.copy2(input_path, temp_file)
                except PermissionError:
                    raise RuntimeError(
                        f"Permission denied when copying file: {input_path}\n\n"
                        f"Possible causes:\n"
                        f"- File is open in another application\n"
                        f"- Insufficient permissions to read the file\n"
                        f"- File is located in a protected system folder"
                    )
                except OSError as e:
                    if "read-only" in str(e).lower():
                        raise RuntimeError(
                            f"Cannot copy file - it may be read-only or in use: {input_path}\n\n"
                            f"Please close any applications using this file and try again."
                        )
                    raise RuntimeError(
                        f"Failed to copy input file to temporary directory: {e}"
                    )

                # Run conversion on temp folder with progress tracking
                try:
                    report = convert_folder_with_progress(
                        input_dir=temp_dir,
                        output_dir=output_path,
                        recursive=recursive,
                        include_doc=include_doc,
                        overwrite=overwrite,
                        progress_callback=self._update_progress,
                        stop_event=self._stop_conversion,
                    )
                finally:
                    # Cleanup temp dir
                    try:
                        shutil.rmtree(temp_dir)
                    except Exception:
                        pass
            else:
                # Folder conversion with progress tracking
                report = convert_folder_with_progress(
                    input_dir=input_path,
                    output_dir=output_path,
                    recursive=recursive,
                    include_doc=include_doc,
                    overwrite=overwrite,
                    progress_callback=self._update_progress,
                    stop_event=self._stop_conversion,
                )

            # Update UI with results
            self._root.after(0, lambda: self._on_conversion_complete(report, None))

        except ValueError as e:
            # Input validation errors from converter
            self._root.after(0, lambda: self._on_conversion_complete(None, e))
        except PermissionError as e:
            # Permission-related errors
            self._root.after(0, lambda: self._on_conversion_complete(
                None,
                RuntimeError(
                    f"Permission denied during conversion.\n\n"
                    f"Error details: {e}\n\n"
                    f"Possible solutions:\n"
                    f"- Close any open Word documents\n"
                    f"- Run the application with administrator privileges\n"
                    f"- Check file/folder permissions\n"
                    f"- Ensure no other program is using the files"
                )
            ))
        except RuntimeError as e:
            # Runtime errors (e.g., LibreOffice not found)
            error_msg = str(e)
            if "LibreOffice" in error_msg or "soffice" in error_msg:
                self._root.after(0, lambda: self._on_conversion_complete(
                    None,
                    RuntimeError(
                        f"LibreOffice conversion failed.\n\n"
                        f"Error: {error_msg}\n\n"
                        f"Possible solutions:\n"
                        f"- Install LibreOffice: https://www.libreoffice.org/download/\n"
                        f"- Ensure 'soffice' is in your system PATH\n"
                        f"- Disable 'Include .doc files' option if you only have .docx files\n"
                        f"- Restart the application after installing LibreOffice"
                    )
                ))
            else:
                self._root.after(0, lambda: self._on_conversion_complete(None, e))
        except OSError as e:
            # OS-level errors (disk full, file in use, etc.)
            error_msg = str(e).lower()
            if "disk full" in error_msg or "no space" in error_msg:
                self._root.after(0, lambda: self._on_conversion_complete(
                    None,
                    RuntimeError(
                        f"Disk full - not enough space to complete conversion.\n\n"
                        f"Please free up disk space and try again."
                    )
                ))
            elif "file in use" in error_msg or "being used" in error_msg:
                self._root.after(0, lambda: self._on_conversion_complete(
                    None,
                    RuntimeError(
                        f"File is in use by another application.\n\n"
                        f"Please close any applications using the files and try again."
                    )
                ))
            else:
                self._root.after(0, lambda: self._on_conversion_complete(
                    None,
                    RuntimeError(f"System error during conversion: {e}")
                ))
        except Exception as e:
            # Unexpected errors
            self._root.after(0, lambda: self._on_conversion_complete(
                None,
                RuntimeError(
                    f"Unexpected error during conversion.\n\n"
                    f"Error: {type(e).__name__}: {e}\n\n"
                    f"Please try again or report this issue if it persists."
                )
            ))

    def _update_progress(self, current: int, total: int, current_file: Path) -> None:
        """
        Update progress bar and status label.

        Args:
            current: Current file index (0-based).
            total: Total number of files.
            current_file: Path to current file being processed.
        """
        # Calculate progress percentage
        if total > 0:
            progress = current / total
        else:
            progress = 0.0

        # Update status text
        if current_file and current_file.name:
            status_text = f"Converting: {current_file.name} ({current + 1}/{total})"
        else:
            status_text = "Conversion complete"

        # Schedule UI updates on main thread
        self._root.after(0, lambda: self._progress_bar.set(progress))
        self._root.after(0, lambda: self._status_label.configure(
            text=status_text,
            text_color="blue",
        ))

    def _on_conversion_complete(self, report: ConversionReport | None, error: Exception | None) -> None:
        """Handle conversion completion."""
        # Update UI state
        self._convert_btn.configure(state="normal")
        self._cancel_btn.configure(state="disabled")
        self._progress_bar.set(1)

        # Check if conversion was cancelled
        if self._stop_conversion.is_set():
            self._status_label.configure(text="Conversion cancelled", text_color="orange")
            self._results_text.configure(state="normal")
            self._results_text.delete("1.0", "end")

            if report:
                result_text = f"Conversion cancelled by user.\n\n"
                result_text += f"Converted: {report.converted}\n"
                result_text += f"Skipped: {report.skipped}\n"
                result_text += f"Failed: {report.failed}\n"
                result_text += "\nPartial results have been preserved."

                if report.failures:
                    result_text += "\n\nFailures:\n"
                    for failure in report.failures:
                        result_text += f"  - {failure}\n"

                self._results_text.insert("1.0", result_text)

                # Update files list with partially converted markdown files
                if self._last_output_path:
                    self._update_files_list(self._last_output_path)

                messagebox.showinfo(
                    "Cancelled",
                    f"Conversion cancelled.\n\nConverted: {report.converted}\nSkipped: {report.skipped}\nFailed: {report.failed}\n\nPartial results have been preserved.",
                )
            else:
                self._results_text.insert("1.0", "Conversion cancelled by user.\nNo files were converted.")
                messagebox.showinfo("Cancelled", "Conversion cancelled by user.\nNo files were converted.")

            self._results_text.configure(state="disabled")
            self._stop_conversion.clear()
            return

        if error:
            self._status_label.configure(text="Conversion failed", text_color="red")
            self._results_text.configure(state="normal")
            self._results_text.delete("1.0", "end")
            self._results_text.insert("1.0", f"Error: {error}\n")
            self._results_text.configure(state="disabled")
            messagebox.showerror("Error", f"Conversion failed: {error}")
            return

        # Update files list with converted markdown files
        if self._last_output_path:
            self._update_files_list(self._last_output_path)

        # Display results
        self._status_label.configure(text="Conversion complete", text_color="green")
        self._results_text.configure(state="normal")
        self._results_text.delete("1.0", "end")

        result_text = f"Conversion completed successfully!\n\n"
        result_text += f"Converted: {report.converted}\n"
        result_text += f"Skipped: {report.skipped}\n"
        result_text += f"Failed: {report.failed}\n"

        if report.failures:
            result_text += "\nFailures:\n"
            for failure in report.failures:
                result_text += f"  - {failure}\n"

        self._results_text.insert("1.0", result_text)
        self._results_text.configure(state="disabled")

        messagebox.showinfo(
            "Success",
            f"Conversion complete!\n\nConverted: {report.converted}\nSkipped: {report.skipped}\nFailed: {report.failed}",
        )

    def _cancel_conversion(self) -> None:
        """Cancel the running conversion."""
        if self._conversion_thread and self._conversion_thread.is_alive():
            # Signal the conversion thread to stop
            self._stop_conversion.set()

            # Update UI to show cancellation in progress
            self._status_label.configure(text="Cancelling conversion...", text_color="orange")
            self._cancel_btn.configure(state="disabled")
            self._results_text.configure(state="normal")
            self._results_text.delete("1.0", "end")
            self._results_text.insert("1.0", "Cancelling conversion...\n")
            self._results_text.configure(state="disabled")
        else:
            # No conversion running
            messagebox.showinfo("Info", "No conversion is currently running.")

    def run(self) -> None:
        """Start the GUI main loop."""
        self._root.mainloop()


def main(argv: list[str] | None = None) -> int:
    """
    Main entry point for the GUI application.

    Args:
        argv: Command line arguments (ignored, for compatibility with CLI).

    Returns:
        Exit code (0 for success, non-zero for failure).
    """
    try:
        app = DocsToMarkdownGUI()
        app.run()
        return 0
    except Exception as e:
        messagebox.showerror(
            "Error",
            f"Failed to start GUI: {e}",
        )
        return 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
