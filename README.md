# docconv

轻量文档格式转换工具：在 **docx / pdf / txt / md / csv / xlsx / html** 之间互转。

纯 Python 实现，零云服务、本地离线可用。带 **图形界面（拖拽批量转换）**，也可当命令行 / Python 库使用。
需要高质量 `docx → pdf` 时，会自动调用本机已安装的 **LibreOffice**（headless），没有也不影响其他转换。

## 特性

- 支持 7 种常见格式互转（见下方支持矩阵）
- 三种用法：图形界面（拖拽批量）、命令行、Python 库
- 本地离线运行，文档不出本机
- `docx → pdf` 借助 LibreOffice 保证排版质量
- 中文友好：`txt / md / csv / xlsx → pdf` 自动启用系统中文字体

## 支持矩阵

| 源 \ 目标 | docx | pdf | txt | md | csv | xlsx | html |
|-----------|:----:|:---:|:---:|:--:|:---:|:----:|:----:|
| **docx**  |  -   |  ✅  |  ✅  | ✅ |  ❌  |  ❌   |  ✅   |
| **pdf**   |  ✅  |  -  |  ❌  | ❌ |  ❌  |  ❌   |  ❌   |
| **txt**   |  ✅  |  ✅  |  -  | ❌ |  ❌  |  ❌   |  ✅   |
| **md**    |  ✅  |  ✅  |  ❌  |  - |  ❌  |  ❌   |  ✅   |
| **csv**   |  ✅  |  ✅  |  ❌  | ❌ |  -   |  ✅   |  ❌   |
| **xlsx**  |  ✅  |  ✅  |  ❌  | ❌ |  ✅  |  -    |  ❌   |
| **html**  |  ❌  |  ❌  |  ✅  | ✅ |  ❌  |  ❌   |  -    |

> 说明：`pdf → txt/md/csv/xlsx` 暂未实现（PDF 文本提取后结构化还原成本较高，后续版本规划）。
> `html → docx/xlsx/pdf` 也未实现（html 解析质量参差，保持「稳」优先）。

## 安装

```bash
git clone https://github.com/zhjade5-svg/docconv.git
cd docconv
pip install -e .

# 或仅安装依赖后直接用模块运行
pip install -r requirements.txt
```

可选：安装 [LibreOffice](https://www.libreoffice.org/) 以获得高质量的 `docx → pdf` 转换。
Windows 常见安装路径会被自动识别；其他平台确保 `soffice` / `libreoffice` 在 PATH 中。

## 快速开始（图形界面）

```bash
# 启动 GUI（拖拽文件 / 批量转换）
python run_gui.py
```

界面操作：把文件拖进窗口（或点「添加文件」）→ 选目标格式 → 「开始转换」。
输出文件默认生成在源文件同目录、同名不同后缀。

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

convert("report.docx", "report.pdf")   # docx -> pdf（需要 LibreOffice）
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
- 未安装 LibreOffice 时，`docx → pdf` 不可用（其他方向不受影响）。
- 转换在源文件同目录生成新文件，不会覆盖源文件（除非显式指定同名不同后缀的输出）。

## 贡献

欢迎提 Issue / PR。新增格式转换请在 `docconv/core.py` 的 `_ROUTES` 中登记路由，
并补充 `tests/` 下的用例。

## 许可证

[MIT](LICENSE) © Jade (zhjade5-svg)
