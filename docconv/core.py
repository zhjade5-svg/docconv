"""docconv 核心转换逻辑。

支持的格式: docx, pdf, txt, md
依赖: pdfplumber, python-docx, reportlab
可选后端: LibreOffice (高质量 docx -> pdf)
"""
from __future__ import annotations

import os
import shutil
import subprocess
from pathlib import Path

import pdfplumber
from docx import Document
from docx.shared import Pt

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

SUPPORTED = {"docx", "pdf", "txt", "md"}


def find_libreoffice() -> "str | None":
    """查找本机 LibreOffice 可执行文件。"""
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return shutil.which("soffice") or shutil.which("libreoffice")


def _get_cjk_font() -> str:
    """尽量注册一个支持中文的字体，失败则回退 Helvetica。"""
    candidates = [
        ("C:/Windows/Fonts/msyh.ttc", 0),
        ("C:/Windows/Fonts/simhei.ttf", 0),
        ("/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc", 0),
        ("/System/Library/Fonts/PingFang.ttc", 0),
    ]
    for path, idx in candidates:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont("CJK", path, subfontIndex=idx))
                return "CJK"
            except Exception:
                continue
    return "Helvetica"


# ------------------------- docx -> pdf -------------------------
def docx_to_pdf(src: str, dst: str) -> None:
    lo = find_libreoffice()
    if not lo:
        raise RuntimeError(
            "未找到 LibreOffice，无法将 docx 转为 pdf。\n"
            "请安装 LibreOffice (https://www.libreoffice.org/) 后重试。"
        )
    outdir = str(Path(dst).parent)
    cmd = [lo, "--headless", "--convert-to", "pdf", "--outdir", outdir, str(src)]
    subprocess.run(cmd, check=True, capture_output=True)
    generated = Path(outdir) / (Path(src).stem + ".pdf")
    if generated.resolve() != Path(dst).resolve():
        generated.replace(dst)


# ------------------------- pdf -> docx -------------------------
def pdf_to_docx(src: str, dst: str) -> None:
    doc = Document()
    with pdfplumber.open(src) as pdf:
        for pi, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            text = page.extract_text() or ""
            if text.strip():
                for line in text.split("\n"):
                    doc.add_paragraph(line)
            for table in tables:
                if not table:
                    continue
                ncols = len(table[0]) if table else 0
                t = doc.add_table(rows=len(table), cols=ncols)
                try:
                    t.style = "Table Grid"
                except Exception:
                    pass
                for r, row in enumerate(table):
                    for c, cell in enumerate(row):
                        t.cell(r, c).text = (cell or "").strip()
            if pi < len(pdf.pages) - 1:
                doc.add_page_break()
    doc.save(dst)


# ------------------------- docx -> txt / md -------------------------
def docx_to_text(src: str, dst: str) -> None:
    doc = Document(src)
    lines = [p.text for p in doc.paragraphs]
    Path(dst).write_text("\n".join(lines), encoding="utf-8")


def docx_to_markdown(src: str, dst: str) -> None:
    doc = Document(src)
    lines = []
    for p in doc.paragraphs:
        style = (p.style.name or "") if p.style else ""
        text = p.text
        if style == "Title":
            lines.append(f"# {text}")
        elif style.startswith("Heading"):
            try:
                level = int(style[len("Heading"):])
            except ValueError:
                level = 1
            lines.append(f"{'#' * min(level, 6)} {text}")
        else:
            lines.append(text)
    Path(dst).write_text("\n".join(lines), encoding="utf-8")


# ------------------------- txt / md -> docx -------------------------
def _split_md(text: str):
    """极简 markdown 解析，返回 (kind, level, content) 列表。"""
    blocks = []
    for raw in text.split("\n"):
        line = raw.rstrip()
        if not line.strip():
            continue
        if line.lstrip().startswith("#"):
            hashes = line.lstrip().split(" ")[0]
            level = len(hashes)
            blocks.append(("heading", min(level, 6), line.lstrip("#").strip()))
        else:
            blocks.append(("p", 0, line))
    return blocks


def text_to_docx(src: str, dst: str) -> None:
    doc = Document()
    text = Path(src).read_text(encoding="utf-8")
    for line in text.split("\n"):
        if line.strip():
            doc.add_paragraph(line)
    doc.save(dst)


def markdown_to_docx(src: str, dst: str) -> None:
    doc = Document()
    text = Path(src).read_text(encoding="utf-8")
    for kind, level, content in _split_md(text):
        if kind == "heading":
            doc.add_heading(content, level=level)
        else:
            doc.add_paragraph(content)
    doc.save(dst)


# ------------------------- txt / md -> pdf -------------------------
def text_to_pdf(src: str, dst: str) -> None:
    font = _get_cjk_font()
    c = canvas.Canvas(str(dst), pagesize=A4)
    width, height = A4
    left = 20 * mm
    top = height - 20 * mm
    line_h = 6 * mm
    c.setFont(font, 11)
    y = top
    for line in Path(src).read_text(encoding="utf-8").split("\n"):
        if y < 20 * mm:
            c.showPage()
            y = top
            c.setFont(font, 11)
        c.drawString(left, y, line[:120])
        y -= line_h
    c.save()


def markdown_to_pdf(src: str, dst: str) -> None:
    text = Path(src).read_text(encoding="utf-8")
    cleaned = "\n".join(
        (ln.lstrip("#").strip() if ln.lstrip().startswith("#") else ln)
        for ln in text.split("\n")
    )
    tmp = Path(src).with_suffix(".txt")
    tmp.write_text(cleaned, encoding="utf-8")
    try:
        text_to_pdf(str(tmp), dst)
    finally:
        try:
            tmp.unlink()
        except OSError:
            pass


# ------------------------- 路由 -------------------------
_ROUTES = {
    ("docx", "pdf"): docx_to_pdf,
    ("pdf", "docx"): pdf_to_docx,
    ("docx", "txt"): docx_to_text,
    ("docx", "md"): docx_to_markdown,
    ("txt", "docx"): text_to_docx,
    ("md", "docx"): markdown_to_docx,
    ("txt", "pdf"): text_to_pdf,
    ("md", "pdf"): markdown_to_pdf,
}


def _ext(path: str) -> str:
    return Path(path).suffix.lstrip(".").lower()


def convert(src: str, dst: str) -> str:
    """把 src 转换为 dst。两者都需带受支持的扩展名。"""
    if not os.path.exists(src):
        raise FileNotFoundError(f"源文件不存在: {src}")
    src_fmt = _ext(src)
    dst_fmt = _ext(dst)
    if src_fmt not in SUPPORTED:
        raise ValueError(f"不支持的源格式: .{src_fmt}")
    if dst_fmt not in SUPPORTED:
        raise ValueError(f"不支持的目标格式: .{dst_fmt}")
    func = _ROUTES.get((src_fmt, dst_fmt))
    if not func:
        raise NotImplementedError(f"暂不支持 {src_fmt} -> {dst_fmt}")
    func(src, dst)
    return dst
