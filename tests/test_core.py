"""docconv 基础测试。"""
import os
import tempfile
from pathlib import Path

from docconv.core import text_to_docx, docx_to_text, docx_to_markdown, markdown_to_docx


def test_text_docx_roundtrip():
    with tempfile.TemporaryDirectory() as d:
        src = Path(d) / "in.txt"
        src.write_text("hello\nworld", encoding="utf-8")
        mid = Path(d) / "mid.docx"
        out = Path(d) / "out.txt"
        text_to_docx(str(src), str(mid))
        docx_to_text(str(mid), str(out))
        assert out.read_text(encoding="utf-8") == "hello\nworld"


def test_markdown_heading_to_docx():
    with tempfile.TemporaryDirectory() as d:
        src = Path(d) / "in.md"
        src.write_text("# 标题\n正文", encoding="utf-8")
        dst = Path(d) / "out.docx"
        markdown_to_docx(str(src), str(dst))
        assert dst.exists()


def test_docx_markdown_roundtrip():
    with tempfile.TemporaryDirectory() as d:
        src = Path(d) / "in.md"
        src.write_text("# 一级\n## 二级\n内容", encoding="utf-8")
        docx = Path(d) / "mid.docx"
        md = Path(d) / "out.md"
        markdown_to_docx(str(src), str(docx))
        docx_to_markdown(str(docx), str(md))
        text = md.read_text(encoding="utf-8")
        assert "# 一级" in text and "## 二级" in text


if __name__ == "__main__":
    test_text_docx_roundtrip()
    test_markdown_heading_to_docx()
    test_docx_markdown_roundtrip()
    print("all tests passed")
