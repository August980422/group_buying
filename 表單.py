#!/usr/bin/env python3
# coding: utf-8
"""
FB group-buy order-sheet generator · v2.6
  • Supports automatic accumulation for the same customer when using the pattern “String+Number”.
  • Global font family / size can be changed via constants at the top of the file.
"""

import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.font as tkfont           # ← new import
from collections import defaultdict, OrderedDict
import pandas as pd

# ── Global font settings ─────────────────────────────────────────────
DEFAULT_FONT_FAMILY = "Segoe UI"        # Change to "Microsoft JhengHei" for Chinese UI, etc.
DEFAULT_FONT_SIZE   = 14                # Adjust font size here
APP_FONT = (DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE)

# ── Regular expressions ─────────────────────────────────────────────
RE_NAME_PLUS = re.compile(r'^([\u4e00-\u9fa5\u00C0-\u024F A-Za-z\s·\._-]+?)\s+\+\s*([0-9]+)\s*$')
RE_NAME_LINE = re.compile(r'^[\u4e00-\u9fa5\u00C0-\u024F A-Za-z\s·\._-]+$')
RE_PURE_NUM    = re.compile(r'^\s*\+?\s*([0-9]+)\s*$')
RE_ANY_LETTER_PLUS = re.compile(r'[A-Za-z]\s*\+\s*\d+')
RE_TIME        = re.compile(r'^\d+\s*(天|日|週|周|小時|分鐘)$')
NOISE_WORDS    = {"回覆", "翻譯年糕"}
RE_STR_QTY     = re.compile(r'^(.*?)(?:\s*)\+\s*(\d+)$')        # Generic “string+qty”

# ── Sort keys ───────────────────────────────────────────────────────
def buyer_key(name: str):
    first = name.lstrip()[0]
    return (0, name.casefold()) if first.isascii() and first.isalpha() else (1, name)

def item_key(item: str):
    first = item.lstrip()[0]
    return (0, item.casefold()) if first.isascii() and first.isalpha() else (1, item)

# ── Parsing logic ───────────────────────────────────────────────────
def parse_orders(text: str) -> dict[str, object]:
    buyers = defaultdict(lambda: None)      # type: ignore
    current: str | None = None              # remember current buyer name

    for raw in text.splitlines():
        line = raw.strip()

        # ── Noise filtering ──
        if (not line or "已編輯" in line or line in NOISE_WORDS or RE_TIME.match(line)):
            continue

        # Same-line “Name +2”
        m = RE_NAME_PLUS.match(line)
        if m:
            _merge_val(buyers, m.group(1).strip(), int(m.group(2)))
            current = None
            continue

        # New buyer’s name line
        if RE_NAME_LINE.match(line):
            current = line
            continue

        # Quantity / combo line (may be multiple lines)
        if current is not None:
            _store_qty(buyers, current, line)

    return buyers

def _store_qty(buyers, name, line):
    if RE_PURE_NUM.match(line) and not RE_ANY_LETTER_PLUS.search(line):
        _merge_val(buyers, name, int(RE_PURE_NUM.match(line).group(1)))   # type: ignore
    else:
        _merge_val(buyers, name, line)

def _merge_val(dic, key, new):
    """Merge `new` into dic[key]; auto-sum if both are numeric or share same prefix."""
    old = dic.get(key)

    # Two original cases remain the same
    if old in (None, ''):
        dic[key] = new
        return
    if isinstance(old, int) and isinstance(new, int):
        dic[key] = old + new
        return

    # New: handle “same prefix string+number” accumulation
    s_old, s_new = str(old).strip(), str(new).strip()
    m_old, m_new = RE_STR_QTY.match(s_old), RE_STR_QTY.match(s_new)

    if m_old and m_new and m_old.group(1) == m_new.group(1):
        qty_sum = int(m_old.group(2)) + int(m_new.group(2))
        dic[key] = f"{m_old.group(1)}+{qty_sum}"
    else:
        # Different prefixes → deduplicate and join with space
        dic[key] = " ".join(sorted({s_old, s_new}))

# ── GUI application ─────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("團購訂單表產生器")
        self.geometry("920x580")

        # Apply global font to all Tk / ttk widgets
        self.option_add("*Font", APP_FONT)
        self._tune_tree_style()

        # {buyer: {item: qty/str}}
        self.data = defaultdict(lambda: OrderedDict())
        self.df_display: pd.DataFrame | None = None
        self._build_ui()

    # Customise ttk.Treeview fonts and row height
    def _tune_tree_style(self):
        style = ttk.Style(self)
        style.configure("Treeview",
                        font=APP_FONT,
                        rowheight=int(DEFAULT_FONT_SIZE * 1.8))
        style.configure("Treeview.Heading",
                        font=(DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE, "bold"))

    # Build all widgets
    def _build_ui(self):
        top = ttk.Frame(self)
        top.pack(fill="x", padx=8, pady=4)

        ttk.Label(top, text="商品名稱：").grid(row=0, column=0, sticky="w")
        self.ent_item = ttk.Entry(top, width=26)
        self.ent_item.grid(row=0, column=1, sticky="w", padx=4)

        ttk.Label(top, text="貼文內容：").grid(row=1, column=0, columnspan=2, sticky="w")
        self.txt_post = tk.Text(top, width=80, height=10, wrap="word")
        self.txt_post.grid(row=2, column=0, columnspan=2, sticky="we", pady=(0, 4))

        btns = ttk.Frame(top)
        btns.grid(row=3, column=0, columnspan=2, sticky="we")
        ttk.Button(btns, text="新增商品", command=self.on_add).pack(side="left", padx=4)
        ttk.Button(btns, text="匯出 Excel", command=self.on_export).pack(side="left", padx=4)

        tblfrm = ttk.Frame(self)
        tblfrm.pack(fill="both", expand=True, padx=8, pady=4)
        self.tree = ttk.Treeview(tblfrm, show="headings")
        vsb = ttk.Scrollbar(tblfrm, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tblfrm, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

    # ── Callbacks ───────────────────────────────────────────────────
    def on_add(self):
        item = self.ent_item.get().strip()
        if not item:
            messagebox.showwarning("缺少商品名稱", "請輸入『商品名稱』")
            return

        buyers = parse_orders(self.txt_post.get("1.0", "end"))
        if not buyers:
            messagebox.showwarning("無符合格式", "找不到任何有效訂購資訊")
            return

        for buyer, val in buyers.items():
            _merge_val(self.data[buyer], item, val)

        self.ent_item.delete(0, "end")
        self.txt_post.delete("1.0", "end")
        self._refresh_tree()

    def on_export(self):
        if self.df_display is None or self.df_display.empty:
            messagebox.showwarning("無資料", "目前沒有可匯出的資料")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel 活頁簿", "*.xlsx")])
        if path:
            self._export_df().to_excel(path, index=False)
            messagebox.showinfo("已匯出", f"檔案已儲存至：\n{path}")

    # ── Helpers ────────────────────────────────────────────────────
    def _refresh_tree(self):
        items = sorted({it for od in self.data.values() for it in od}, key=item_key)
        buyers = sorted(self.data.keys(), key=buyer_key)

        rows = [dict(self.data[b]) for b in buyers]
        df = pd.DataFrame(rows, index=buyers).reindex(columns=items).fillna("")
        df.index.name = "姓名"
        df.reset_index(inplace=True)
        self.df_display = df

        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center",
                             width=max(80, int(len(col) * 14)))
        for _, row in df.iterrows():
            self.tree.insert("", "end", values=list(row))

    def _export_df(self) -> pd.DataFrame:
        df = self.df_display.copy()
        df.insert(0, 'Unnamed: 0', '')
        df.rename(columns={'姓名': 'Unnamed: 1'}, inplace=True)
        return df

# ── Entry point ────────────────────────────────────────────────────
if __name__ == "__main__":
    App().mainloop()
