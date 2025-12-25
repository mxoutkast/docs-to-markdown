from pathlib import Path

from docs_to_markdown.converter import _assets_dir_for_md, _iter_input_files, _output_md_path


def test_iter_input_files_non_recursive(tmp_path: Path) -> None:
    (tmp_path / "a.docx").write_text("x")
    (tmp_path / "b.doc").write_text("x")
    (tmp_path / "c.txt").write_text("x")
    sub = tmp_path / "sub"
    sub.mkdir()
    (sub / "d.docx").write_text("x")

    files = _iter_input_files(tmp_path, recursive=False, include_doc=False)
    assert [p.name for p in files] == ["a.docx"]


def test_iter_input_files_recursive_with_doc(tmp_path: Path) -> None:
    (tmp_path / "a.docx").write_text("x")
    (tmp_path / "b.doc").write_text("x")
    sub = tmp_path / "sub"
    sub.mkdir()
    (sub / "d.docx").write_text("x")

    files = _iter_input_files(tmp_path, recursive=True, include_doc=True)
    assert [p.name for p in files] == ["a.docx", "b.doc", "d.docx"]


def test_output_md_path_mirrors_structure(tmp_path: Path) -> None:
    input_dir = tmp_path / "in"
    output_dir = tmp_path / "out"
    (input_dir / "sub").mkdir(parents=True)

    src = input_dir / "sub" / "My File.DOCX"
    md = _output_md_path(src, input_dir=input_dir, output_dir=output_dir)

    assert md == output_dir / "sub" / "My File.md"


def test_assets_dir_for_md_is_sibling(tmp_path: Path) -> None:
    md_path = tmp_path / "out" / "My Doc.md"
    assets = _assets_dir_for_md(md_path)
    assert assets == tmp_path / "out" / "My Doc_files"
