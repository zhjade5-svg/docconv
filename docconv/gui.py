"""docconv 图形界面：拖拽文件批量转换。"""
from __future__ import annotations

import re
import threading
from pathlib import Path
from tkinter import (
    Tk,
    Frame,
    Label,
    Button,
    Listbox,
    Scrollbar,
    StringVar,
    Text,
    END,
    messagebox,
    filedialog,
)
from tkinter.ttk import Combobox

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
except Exception:  # 缺 tkinterdnd2 时降级为普通 Tk，仅失去拖拽
    TkinterDnD = None
    DND_FILES = None

from .core import convert, SUPPORTED

FORMATS = sorted(SUPPORTED)


class App:
    def __init__(self, root):
        self.root = root
        self.files = []
        self.target = StringVar(value="pdf")
        root.title("docconv · 文档格式转换")
        root.geometry("680x500")

        Label(
            root,
            text="把文件拖到这里，或点「添加文件」",
            font=("Microsoft YaHei", 12),
        ).pack(pady=8)

        frm = Frame(root)
        frm.pack(fill="both", expand=True, padx=12)

        self.listbox = Listbox(frm, selectmode="extended")
        self.listbox.pack(side="left", fill="both", expand=True)
        sb = Scrollbar(frm, command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.config(yscrollcommand=sb.set)

        if TkinterDnD is not None:
            try:
                root.drop_target_register(DND_FILES)
                root.dnd_bind("<<Drop>>", self.on_drop)
            except Exception:
                pass

        ctrl = Frame(root)
        ctrl.pack(fill="x", padx=12, pady=6)
        Label(ctrl, text="目标格式:").pack(side="left")
        Combobox(
            ctrl,
            textvariable=self.target,
            values=FORMATS,
            width=10,
            state="readonly",
        ).pack(side="left", padx=6)
        Button(ctrl, text="添加文件", command=self.add_files).pack(side="left", padx=4)
        Button(ctrl, text="移除选中", command=self.remove_sel).pack(side="left", padx=4)
        Button(ctrl, text="清空", command=self.clear).pack(side="left", padx=4)
        Button(ctrl, text="开始转换", command=self.start).pack(side="right", padx=4)

        self.log = Text(root, height=9, state="disabled", bg="#f5f5f5")
        self.log.pack(fill="x", padx=12, pady=6)

    # ---------- 日志 ----------
    def log_msg(self, msg):
        self.log.config(state="normal")
        self.log.insert(END, msg + "\n")
        self.log.see(END)
        self.log.config(state="disabled")

    # ---------- 文件管理 ----------
    def _parse_drop(self, data):
        data = data.strip()
        if "{" in data:
            return re.findall(r"\{([^}]*)\}", data)
        return data.split()

    def on_drop(self, event):
        for f in self._parse_drop(event.data):
            self._add(f)

    def add_files(self):
        paths = filedialog.askopenfilenames(title="选择文件")
        for p in paths:
            self._add(p)

    def _add(self, path):
        p = str(Path(path))
        if p not in self.files:
            self.files.append(p)
            self.listbox.insert(END, p)

    def remove_sel(self):
        for i in reversed(self.listbox.curselection()):
            self.files.pop(i)
            self.listbox.delete(i)

    def clear(self):
        self.files.clear()
        self.listbox.delete(0, END)

    # ---------- 转换 ----------
    def start(self):
        if not self.files:
            messagebox.showwarning("提示", "请先添加文件")
            return
        to = self.target.get()
        threading.Thread(target=self._run, args=(to,), daemon=True).start()

    def _run(self, to):
        self.log_msg(f"开始转换，目标格式: .{to}")
        ok = 0
        for src in list(self.files):
            dst = str(Path(src).with_suffix("." + to))
            try:
                convert(src, dst)
                ok += 1
                self.log_msg(f"✓ {Path(src).name} -> {Path(dst).name}")
            except Exception as e:
                self.log_msg(f"✗ {Path(src).name}: {e}")
        self.log_msg(f"完成：成功 {ok}/{len(self.files)}")
        messagebox.showinfo("完成", f"转换完成：成功 {ok}/{len(self.files)}")


def main():
    root = TkinterDnD.Tk() if TkinterDnD is not None else Tk()
    App(root)
    root.mainloop()
