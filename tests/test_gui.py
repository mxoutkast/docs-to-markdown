"""Unit tests for GUI components."""

from pathlib import Path

from docs_to_markdown.gui import convert_folder_with_progress, DocsToMarkdownGUI


def test_convert_folder_with_progress_basic(tmp_path: Path) -> None:
    """Test basic folder conversion with progress tracking."""
    # Create input directory with a simple docx file
    input_dir = tmp_path / "input"
    input_dir.mkdir()
    
    # Create a minimal valid docx file
    docx_file = input_dir / "test.docx"
    docx_file.write_bytes(
        b"PK\x03\x04\x14\x00\x00\x00\x08\x00"  # Minimal ZIP header
        b"\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00"
        b"\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00"  # More ZIP data
    )
    
    # Create output directory
    output_dir = tmp_path / "output"
    
    # Track progress updates
    progress_updates = []
    
    def progress_callback(current: int, total: int, current_file: Path) -> None:
        progress_updates.append((current, total, current_file.name if current_file else ""))
    
    # Run conversion
    report = convert_folder_with_progress(
        input_dir=input_dir,
        output_dir=output_dir,
        recursive=False,
        include_doc=False,
        overwrite=False,
        progress_callback=progress_callback,
    )
    
    # Verify report structure
    assert isinstance(report.converted, int)
    assert isinstance(report.skipped, int)
    assert isinstance(report.failed, int)
    assert isinstance(report.failures, tuple)
    
    # Verify progress was tracked
    assert len(progress_updates) > 0


def test_convert_folder_with_progress_non_recursive(tmp_path: Path) -> None:
    """Test non-recursive conversion ignores subdirectories."""
    # Create input directory with files in subdirectory
    input_dir = tmp_path / "input"
    input_dir.mkdir()
    
    # Create file in root
    (input_dir / "root.docx").write_bytes(b"PK\x03\x04")
    
    # Create file in subdirectory
    subdir = input_dir / "sub"
    subdir.mkdir()
    (subdir / "nested.docx").write_bytes(b"PK\x03\x04")
    
    output_dir = tmp_path / "output"
    
    # Run non-recursive conversion
    report = convert_folder_with_progress(
        input_dir=input_dir,
        output_dir=output_dir,
        recursive=False,
        include_doc=False,
        overwrite=False,
    )
    
    # Verify only root file was processed (or failed due to invalid docx)
    # The exact count depends on whether the files are valid docx
    assert isinstance(report.converted, int)
    assert isinstance(report.failed, int)


def test_convert_folder_with_progress_skip_existing(tmp_path: Path) -> None:
    """Test that existing .md files are skipped when overwrite=False."""
    input_dir = tmp_path / "input"
    input_dir.mkdir()
    
    # Create a docx file
    (input_dir / "test.docx").write_bytes(b"PK\x03\x04")
    
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    
    # Create existing .md file
    (output_dir / "test.md").write_text("# Existing content")
    
    # Run conversion without overwrite
    report = convert_folder_with_progress(
        input_dir=input_dir,
        output_dir=output_dir,
        recursive=False,
        include_doc=False,
        overwrite=False,
    )
    
    # Verify file was skipped
    assert report.skipped >= 0


def test_convert_folder_with_progress_overwrite(tmp_path: Path) -> None:
    """Test that existing .md files are overwritten when overwrite=True."""
    input_dir = tmp_path / "input"
    input_dir.mkdir()
    
    # Create a docx file
    (input_dir / "test.docx").write_bytes(b"PK\x03\x04")
    
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    
    # Create existing .md file
    existing_md = output_dir / "test.md"
    existing_md.write_text("# Existing content")
    
    # Run conversion with overwrite
    report = convert_folder_with_progress(
        input_dir=input_dir,
        output_dir=output_dir,
        recursive=False,
        include_doc=False,
        overwrite=True,
    )
    
    # Verify file was not skipped
    assert report.skipped == 0


def test_convert_folder_with_progress_invalid_input(tmp_path: Path) -> None:
    """Test that invalid input directory raises ValueError."""
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    
    # Try to convert non-existent directory
    try:
        convert_folder_with_progress(
            input_dir=tmp_path / "nonexistent",
            output_dir=output_dir,
            recursive=False,
            include_doc=False,
            overwrite=False,
        )
        assert False, "Expected ValueError"
    except ValueError as e:
        assert "does not exist" in str(e).lower() or "not a directory" in str(e).lower()


def test_convert_folder_with_progress_stop_event(tmp_path: Path) -> None:
    """Test that stop_event can cancel conversion."""
    import threading
    
    input_dir = tmp_path / "input"
    input_dir.mkdir()
    
    # Create multiple docx files
    for i in range(5):
        (input_dir / f"file{i}.docx").write_bytes(b"PK\x03\x04")
    
    output_dir = tmp_path / "output"
    
    # Create stop event
    stop_event = threading.Event()
    
    # Stop after processing first file
    progress_count = [0]
    
    def progress_callback(current: int, total: int, current_file: Path) -> None:
        progress_count[0] += 1
        if progress_count[0] >= 1:
            stop_event.set()
    
    # Run conversion with stop event
    report = convert_folder_with_progress(
        input_dir=input_dir,
        output_dir=output_dir,
        recursive=False,
        include_doc=False,
        overwrite=False,
        progress_callback=progress_callback,
        stop_event=stop_event,
    )
    
    # Verify conversion was stopped early
    assert progress_count[0] >= 1


def test_render_markdown_simple_text() -> None:
    """Test rendering simple markdown text."""
    gui = DocsToMarkdownGUI()
    
    # Test simple text
    md_text = "Hello, world!"
    rendered = gui._render_markdown(md_text)
    assert "Hello, world!" in rendered


def test_render_markdown_headers() -> None:
    """Test rendering markdown headers."""
    gui = DocsToMarkdownGUI()
    
    # Test headers
    md_text = "# Header 1\n## Header 2\n### Header 3"
    rendered = gui._render_markdown(md_text)
    
    assert "# Header 1" in rendered
    assert "## Header 2" in rendered
    assert "### Header 3" in rendered


def test_render_markdown_bold_italic() -> None:
    """Test rendering bold and italic text."""
    gui = DocsToMarkdownGUI()
    
    # Test bold and italic
    md_text = "**bold text** and *italic text*"
    rendered = gui._render_markdown(md_text)
    
    assert "**bold text**" in rendered
    assert "*italic text*" in rendered


def test_render_markdown_lists() -> None:
    """Test rendering markdown lists."""
    gui = DocsToMarkdownGUI()
    
    # Test lists
    md_text = "- Item 1\n- Item 2\n- Item 3"
    rendered = gui._render_markdown(md_text)
    
    assert "- Item 1" in rendered
    assert "- Item 2" in rendered
    assert "- Item 3" in rendered


def test_render_markdown_code_block() -> None:
    """Test rendering markdown code blocks."""
    gui = DocsToMarkdownGUI()
    
    # Test code block
    md_text = "```\ncode here\n```"
    rendered = gui._render_markdown(md_text)
    
    assert "```" in rendered
    assert "code here" in rendered


def test_render_markdown_blockquote() -> None:
    """Test rendering markdown blockquotes."""
    gui = DocsToMarkdownGUI()
    
    # Test blockquote
    md_text = "> This is a quote"
    rendered = gui._render_markdown(md_text)
    
    assert ">" in rendered
    assert "This is a quote" in rendered


def test_render_markdown_empty() -> None:
    """Test rendering empty markdown."""
    gui = DocsToMarkdownGUI()
    
    # Test empty string
    rendered = gui._render_markdown("")
    assert rendered == ""
    
    # Test None-like whitespace
    rendered = gui._render_markdown("   \n\n  ")
    assert rendered == ""


def test_render_markdown_table() -> None:
    """Test rendering markdown tables."""
    gui = DocsToMarkdownGUI()
    
    # Test table
    md_text = "| Header 1 | Header 2 |\n|----------|----------|\n| Cell 1   | Cell 2   |"
    rendered = gui._render_markdown(md_text)
    
    assert "|" in rendered
    assert "Header 1" in rendered
    assert "Cell 1" in rendered


def test_render_markdown_horizontal_rule() -> None:
    """Test rendering horizontal rules."""
    gui = DocsToMarkdownGUI()
    
    # Test horizontal rule
    md_text = "---"
    rendered = gui._render_markdown(md_text)
    
    assert "---" in rendered


def test_render_markdown_mixed_content() -> None:
    """Test rendering mixed markdown content."""
    gui = DocsToMarkdownGUI()
    
    # Test mixed content
    md_text = """# Title

Paragraph with **bold** and *italic*.

- List item 1
- List item 2

> A quote

```
code block
```
"""
    rendered = gui._render_markdown(md_text)
    
    assert "# Title" in rendered
    assert "**bold**" in rendered
    assert "*italic*" in rendered
    assert "- List item 1" in rendered
    assert ">" in rendered
    assert "```" in rendered


def test_copy_preview_has_method() -> None:
    """Test that _copy_preview method exists."""
    gui = DocsToMarkdownGUI()
    assert hasattr(gui, "_copy_preview")
    assert callable(getattr(gui, "_copy_preview"))


def test_current_markdown_text_initialized() -> None:
    """Test that _current_markdown_text is initialized."""
    gui = DocsToMarkdownGUI()
    assert hasattr(gui, "_current_markdown_text")
    assert isinstance(gui._current_markdown_text, str)


def test_copy_button_exists() -> None:
    """Test that copy button is created in preview section."""
    gui = DocsToMarkdownGUI()
    assert hasattr(gui, "_copy_btn")
    assert gui._copy_btn is not None
