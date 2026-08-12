# docconv

轻量文档格式转换工具：在 **docx / pdf / txt / md** 之间互转。

纯 Python 实现，零云服务、本地离线可用。需要高质量 `docx → pdf` 时，
会自动调用本机已安装的 **LibreOffice**（headless），没有也不影响其他转换。

## 特性

- 支持多种格式互转（见下方支持矩阵）
- 命令行一键转换，也能当 Python 库调用
- 本地离线运行，文档不出本机
- `docx → pdf` 借助 LibreOffice 保证排版质量
- 中文友好：`txt / md → pdf` 自动启用系统中文字体

## 支持矩阵

| 源 \ 目标 | docx | pdf | txt | md |
|-----------|:----:|:---:|:---:|:--:|
| **docx**  |  -   |  ✅  |  ✅  | ✅ |
| **pdf**   |  ✅  |  -  |  ❌  | ❌ |
| **txt**   |  ✅  |  ✅  |  -  | ❌ |
| **md**    |  ✅  |  ✅  |  ❌  |  - |

> 说明：`pdf → txt/md` 暂未实现（PDF 文本提取后结构化还原成本较高，后续版本规划）。

## 安装

```bash
# 从源码安装
git clone https://github.com/zhjade5-svg/docconv.git
cd docconv
pip install -e .

# 或仅安装依赖后直接用模块运行
pip install -r requirements.txt
```

可选：安装 [LibreOffice](https://www.libreoffice.org/) 以获得高质量的 `docx → pdf` 转换。
Windows 常见安装路径会被自动识别；其他平台确保 `soffice` / `libreoffice` 在 PATH 中。

## 快速开始（命令行）

```bash
# 转成 pdf（省略输出名时，默认同目录同名 .pdf）
docconv convert report.docx

# 指定输出
docconv convert report.docx output.pdf

# 用 --to 指定目标格式
docconv convert report.docx --to txt

# 列出支持的转换
docconv formats
```

## 作为 Python 库使用

```python
from docconv import convert

# docx -> pdf（需要 LibreOffice）
convert("report.docx", "report.pdf")

# txt -> docx
convert("note.txt", "note.docx")

# md -> pdf
convert("readme.md", "readme.pdf")
```

## 依赖

- `pdfplumber` — 读取 PDF、提取文本与表格
- `python-docx` — 读写 Word (.docx)
- `reportlab` — 将文本 / Markdown 排成 PDF

## 已知限制

- `pdf → docx` 为提取式转换：文本按段落还原，表格以表格形式追加；
  PDF 中图片暂不直接保留。
- `txt / md → pdf` 为简单排版（等宽/逐行），不含复杂样式。
- 未安装 LibreOffice 时，`docx → pdf` 不可用（其他方向不受影响）。

## 贡献

欢迎提 Issue / PR。新增格式转换请在 `docconv/core.py` 的 `_ROUTES` 中登记路由，
并补充 `tests/` 下的用例。

## 许可证

[MIT](LICENSE) © Jade (zhjade5-svg)
