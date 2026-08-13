# docconv

轻量文档格式转换工具：在 **doc / docx / pdf / txt / md / csv / xlsx / html** 之间互转。

纯 Python 实现，零云服务、本地离线可用。带 **图形界面（拖拽批量转换）**，也可当命令行 / Python 库使用。
用到 `.doc` 或高质量 `docx → pdf` 时，exe 会**首次自动下载并缓存便携 LibreOffice（约 350MB，仅一次，之后离线可用）**；其余格式完全纯 Python，连 LibreOffice 都不需要。

## 特性

- 支持 8 种常见格式互转（见下方支持矩阵）
- 三种用法：图形界面（拖拽批量）、命令行、Python 库
- 本地离线运行，文档不出本机
- **LibreOffice 按需自动融入**：首次用 `.doc` / 高质量 `docx → pdf` 时自动下载并缓存便携 LO，零手动安装；其余格式纯 Python，不依赖 LibreOffice
- 中文友好：`txt / md / csv / xlsx → pdf` 自动启用系统中文字体

## 支持矩阵

| 源 \ 目标 | doc | docx | pdf | txt | md | csv | xlsx | html |
|-----------|:---:|:----:|:---:|:---:|:--:|:---:|:----:|:----:|
| **doc**   |  -  |  ✅  |  ✅  | ✅ | ✅ |  ❌  |  ❌   |  ✅   |
| **docx**  |  ✅  |  -   |  ✅  | ✅ | ✅ |  ❌  |  ❌   |  ✅   |
| **pdf**   |  ❌  |  ✅  |  -  | ❌ | ❌ |  ❌  |  ❌   |  ❌   |
| **txt**   |  ❌  |  ✅  |  ✅  |  - | ❌ |  ❌  |  ❌   |  ✅   |
| **md**    |  ❌  |  ✅  |  ✅  | ❌ |  - |  ❌  |  ❌   |  ✅   |
| **csv**   |  ❌  |  ✅  |  ✅  | ❌ | ❌ |  -   |  ✅   |  ❌   |
| **xlsx**  |  ❌  |  ✅  |  ✅  | ❌ | ❌ |  ✅  |  -    |  ❌   |
| **html**  |  ❌  |  ❌  |  ❌  | ✅ | ✅ |  ❌  |  ❌   |  -    |

> 说明：`.doc` 与高质量 `docx → pdf` 依赖 LibreOffice。exe 会在首次使用时**自动下载并缓存便携版**（也可使用系统已装的 LibreOffice，或把 `libreoffice/` 文件夹放在 exe 同级）。其余转换纯 Python，不依赖 LibreOffice。
> `pdf → txt/md/csv/xlsx/doc/html` 暂未实现；`html → docx/xlsx/pdf` 同理未实现（保持「稳」优先）。

## 安装

```bash
git clone https://github.com/zhjade5-svg/docconv.git
cd docconv
pip install -e .

# 或仅安装依赖后直接用模块运行
pip install -r requirements.txt
```

## 快速开始（图形界面）

```bash
# 启动 GUI（拖拽文件 / 批量转换）
python run_gui.py
```

界面操作：把文件拖进窗口（或点「添加文件」）→ 选目标格式 →（可选）设「输出目录」→ 「开始转换」。

- **输出目录**留空 = 输出文件生成在源文件同目录、同名不同后缀；
- 指定输出目录后，所有转换结果落到该目录；目录不存在时自动回退到源文件目录。
- 转换日志会打印每个结果的完整保存路径，方便查找。

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

convert("report.docx", "report.pdf")   # docx -> pdf（优先 LibreOffice 保排版，无则纯 Python）
convert("note.txt", "note.docx")       # txt -> docx
convert("table.csv", "table.xlsx")     # csv -> xlsx
convert("page.html", "page.md")        # html -> md
```

## 打包成独立 exe（发行版）

仓库已提供图形界面，可用 PyInstaller 打包成单文件 Windows 程序：

```bash
pip install pyinstaller tkinterdnd2
pyinstaller --onefile --windowed --collect-all tkinterdnd2 --name docconv run_gui.py
# 产物：dist/docconv.exe
```

`--collect-all tkinterdnd2` 必须带上，否则冻结版拖拽功能失效。

> 发行版默认不含 LibreOffice 本体（约 350MB），改为**首次使用时自动下载并缓存**到本地（零手动安装）。若想完全离线预置，可把一份 LibreOffice 放到 exe 同级的 `libreoffice/` 目录下（即 `libreoffice/program/soffice.exe`），exe 会优先使用它。

## 依赖

- `pdfplumber` — 读取 PDF、提取文本与表格
- `python-docx` — 读写 Word (.docx)
- `reportlab` — 将文本 / Markdown / 表格排成 PDF
- `openpyxl` — 读写 Excel (.xlsx) / 与 CSV 互转
- `html2text` — HTML → 文本 / Markdown
- `markdown` — Markdown → HTML
- `tkinterdnd2` — GUI 拖拽（无界面环境可省略）

## 已知限制

- `pdf → docx` 为提取式转换：文本按段落还原，表格以表格形式追加；PDF 中图片暂不直接保留。
- `txt / md → pdf` 为简单排版（逐行），不含复杂样式。
- 转换在源文件同目录生成新文件，不会覆盖源文件（除非显式指定同名不同后缀的输出）。

## 贡献

欢迎提 Issue / PR。新增格式转换请在 `docconv/core.py` 的 `_ROUTES` 中登记路由，
并补充 `tests/` 下的用例。

## 许可证

[MIT](LICENSE) © Jade (zhjade5-svg)
