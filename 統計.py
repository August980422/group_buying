import re
from collections import defaultdict
import tkinter as tk
from tkinter import ttk

# 創建主視窗
root = tk.Tk()
root.title("叫單統計")

# 設置視窗大小
root.geometry("500x600")

# 定義一個變量來儲存統計結果
order_count = defaultdict(int)


# 定義一個函數來統計叫單數據
def calculate_statistics():
    global order_count  # 使用全局變量來儲存結果
    comments = text_input.get("1.0", tk.END)  # 取得文字輸入框中的內容
    order_count.clear()  # 清空之前的統計數據

    # 使用正則表達式來匹配有無空格的叫單項目  
    orders = [
        (letter.upper(), number)
        for letter, number in re.findall(r"([A-Za-z])\s*[\+\*]?\s*(\d+)", comments)
    ]

    # 統計每個項目的叫單數量
    for order in orders:
        product = order[0]  # 抓取產品名稱，例如 A、B、C
        quantity = int(order[1])  # 抓取對應的數量
        order_count[product] += quantity

    # 清空表格 
    for i in tree.get_children():
        tree.delete(i)

    # 將統計數據插入表格中
    for product, count in sorted(order_count.items()):
        tree.insert("", tk.END, values=(product, count))


# 定義一個函數來複製統計結果到剪貼簿
def copy_to_clipboard():
    result = []
    for product, count in sorted(order_count.items()):
        result.append(f"{product}.{count}")

    # 清空剪貼簿並複製結果
    root.clipboard_clear()
    root.clipboard_append(",".join(result))
    root.update()

def clear_all():
    text_input.delete("1.0", tk.END)
    for i in tree.get_children():
        tree.delete(i)
    order_count.clear()  # 清空統計數據


# 創建文字輸入框，讓使用者輸入留言資料
label = tk.Label(root, text="請輸入留言資料：")
label.pack(pady=5)

# 創建一個框架來包含文字輸入框和捲動條
text_frame = tk.Frame(root)
text_frame.pack(pady=10)

# 創建文字輸入框
text_input = tk.Text(text_frame, height=10, width=50, wrap="word")

# 創建垂直捲動條
scrollbar = tk.Scrollbar(text_frame, command=text_input.yview)
scrollbar.pack(side="right", fill="y")

# 將捲動條與文字框連結
text_input.config(yscrollcommand=scrollbar.set)
text_input.pack(side="left", fill="both")

# 創建一個框架來包含表格和按鈕，確保它們並排顯示
result_frame = tk.Frame(root)
result_frame.pack(pady=10)

# 創建一個表格來顯示結果
tree = ttk.Treeview(
    result_frame, columns=("product", "count"), show="headings", height=15
)
tree.heading("product", text="產品")
tree.heading("count", text="數量")
tree.column("product", width=100, anchor="center")
tree.column("count", width=100, anchor="center")

# 在框架中放置表格
tree.pack(side="left", padx=10)


# 創建一個標籤來顯示複製狀態
copied_label = tk.Label(result_frame, text="")
copied_label.pack(side="left", padx=10)

# 創建一個框架來包含「計算統計」和「複製結果」按鈕
button_frame = tk.Frame(root)
button_frame.pack(pady=10)

# 創建「計算統計」按鈕，並放在框架中
calculate_button = tk.Button(
    button_frame, text="計算統計", command=calculate_statistics
)
calculate_button.pack(side="left", padx=10)

# 創建「複製結果」按鈕，並放在「計算統計」按鈕的旁邊
copy_button_bottom = tk.Button(button_frame, text="複製結果", command=copy_to_clipboard)
copy_button_bottom.pack(side="left", padx=10)

# 創建「清除」按鈕，並放在「複製結果」按鈕的旁邊
clear_button = tk.Button(button_frame, text="清除", command=clear_all)
clear_button.pack(side="left", padx=10)

root.mainloop()