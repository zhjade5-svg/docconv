"""docconv 命令行入口。"""
import argparse
from pathlib import Path

from .core import convert, SUPPORTED, _ROUTES


def _fmt_help() -> str:
    pairs = sorted(set(_ROUTES.keys()))
    return "\n".join(f"  {a} -> {b}" for a, b in pairs)


def main(argv=None) -> None:
    ap = argparse.ArgumentParser(
        prog="docconv",
        description="轻量文档格式转换工具 (docx / pdf / txt / md)",
    )
    sub = ap.add_subparsers(dest="cmd")

    c = sub.add_parser("convert", help="转换单个文件")
    c.add_argument("input", help="源文件路径")
    c.add_argument("output", nargs="?", help="输出路径（含扩展名）")
    c.add_argument(
        "--to",
        help="目标格式: pdf/docx/txt/md（省略 output 时可用）",
    )

    sub.add_parser("formats", help="列出支持的转换")

    args = ap.parse_args(argv)

    if args.cmd == "formats":
        print("支持的转换:")
        print(_fmt_help())
        return

    if args.cmd == "convert":
        src = args.input
        if not args.output:
            to = (args.to or "pdf").lower()
            dst = str(Path(src).with_suffix("." + to))
        else:
            dst = args.output
        out = convert(src, dst)
        print(f"已转换: {src} -> {out}")
        return

    ap.print_help()


if __name__ == "__main__":
    main()
