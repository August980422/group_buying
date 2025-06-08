#!/usr/bin/env python3
# coding: utf-8
"""
Order Counter with dual modes:
  • Mode "product": counts each product code (A+1, B+2 …).
  • Mode "total"  : sums all '+n' regardless of product codes.
Supports full-width / half-width characters via Unicode NFKC normalization.
"""

import re
import unicodedata
from collections import defaultdict
import tkinter as tk
from tkinter import ttk

# ── main window ───────────────────────────────────────────────
root = tk.Tk()
root.title("Order Counter")
root.geometry("500x600")

order_count: dict[str, int] = defaultdict(int)  # per-product counts
total_qty: int = 0                               # grand total for '+n' mode
mode_var = tk.StringVar(value="product")         # default statistics mode

# ── functions ─────────────────────────────────────────────────
def calculate_statistics() -> None:
    """Parse the comment area and refresh the result table."""
    global total_qty
    comments_raw = text_input.get("1.0", tk.END)
    comments = unicodedata.normalize("NFKC", comments_raw)   # full-width → half-width

    order_count.clear()
    total_qty = 0
    tree.delete(*tree.get_children())

    if mode_var.get() == "product":
        # Capture like A+1、b 2、C*3 (case-insensitive)
        for letter, num in re.findall(r"([A-Za-z])\s*[\+\*]?\s*(\d+)", comments):
            order_count[letter.upper()] += int(num)
        for product, count in sorted(order_count.items()):
            tree.insert("", tk.END, values=(product, count))
    else:  # total mode
        numbers = [int(n) for n in re.findall(r"[\+\*]\s*(\d+)", comments)]
        total_qty = sum(numbers)
        tree.insert("", tk.END, values=("TOTAL", total_qty))

def copy_to_clipboard() -> None:
    """Copy statistics to clipboard in a compact form."""
    if mode_var.get() == "product":
        content = ",".join(f"{p}.{c}" for p, c in sorted(order_count.items()))
    else:
        content = str(total_qty)

    root.clipboard_clear()
    root.clipboard_append(content)
    root.update()  # keep clipboard after window closes

def clear_all() -> None:
    """Clear text input, table and stored statistics."""
    text_input.delete("1.0", tk.END)
    tree.delete(*tree.get_children())
    order_count.clear()
    global total_qty
    total_qty = 0

# ── widgets ───────────────────────────────────────────────────
# input label + text box with scrollbar
tk.Label(root, text="請輸入留言資料：").pack(pady=5)
text_frame = tk.Frame(root)
text_frame.pack(pady=10)

text_input = tk.Text(text_frame, height=10, width=50, wrap="word")
scrollbar = tk.Scrollbar(text_frame, command=text_input.yview)
text_input.config(yscrollcommand=scrollbar.set)
text_input.pack(side="left", fill="both")
scrollbar.pack(side="right", fill="y")

# mode selection
mode_frame = tk.Frame(root)
mode_frame.pack(pady=5)
tk.Radiobutton(mode_frame, text="依產品統計 (A+1)", variable=mode_var,
               value="product").pack(side="left", padx=5)
tk.Radiobutton(mode_frame, text="僅計算 +n 總量", variable=mode_var,
               value="total").pack(side="left", padx=5)

# result table
result_frame = tk.Frame(root)
result_frame.pack(pady=10)

tree = ttk.Treeview(result_frame, columns=("product", "count"),
                    show="headings", height=15)
tree.heading("product", text="產品")
tree.heading("count", text="數量")
tree.column("product", width=100, anchor="center")
tree.column("count", width=100, anchor="center")
tree.pack(side="left", padx=10)

# buttons
button_frame = tk.Frame(root)
button_frame.pack(pady=10)

tk.Button(button_frame, text="計算統計",
          command=calculate_statistics).pack(side="left", padx=10)
tk.Button(button_frame, text="複製結果",
          command=copy_to_clipboard).pack(side="left", padx=10)
tk.Button(button_frame, text="清除",
          command=clear_all).pack(side="left", padx=10)

root.mainloop()
