## 🚀 Quick Start

```bash
# Install dependencies
pip install pandas openpyxl customtkinter pyperclip

# Run tools
python 表單.py         # Order Sheet Generator
python 訊息.py         # Order Message Generator
python 統計.py         # Comment Statistics Tool
```

---

## 📁 Tool Descriptions

### 1️⃣ `表單.py` – Order Sheet Generator

**Description:**

A GUI tool designed to parse unstructured group-buy comments and convert them into a structured Excel order sheet. It supports various comment formats and automatically aggregates orders by buyer and product.

**Key Features:**

* Parses formats like `Name + Quantity` or separated lines.
* Supports Chinese, English, and mixed text input.
* Automatically merges duplicate buyers and quantities.
* Supports accumulation of “same-prefix” strings (e.g., `奶茶+1` + `奶茶+2` → `奶茶+3`).
* Displays a sortable and scrollable preview table.
* Exportable to `.xlsx` Excel format.
* Built with `tkinter` for ease of use.

**Usage:**

```bash
python 表單.py
```

1. Enter a product name in the input field.
2. Paste comment content into the text area (e.g., from Facebook/LINE).
3. Click **新增商品** ("Add Product").
4. Add more products if needed.
5. Click **匯出 Excel** ("Export to Excel") to save the result.

---

### 2️⃣ `訊息.py` – Order Message Generator

**Description:**

This tool reads an Excel sheet containing buyer orders and automatically generates detailed per-buyer messages. These messages are formatted for delivery via private chat or messaging platforms and include product details, unit prices, and quantities.

**Key Features:**

* Accepts Excel files with headers: first row for prices, following rows for order data.
* Auto-generates structured messages for each buyer.
* Built with `customtkinter` for modern UI design.
* Supports:

  * Clipboard copying.
  * Text message editing and saving.
  * Filtering buyers by name.
  * Sorting buyers based on ordered products.
  * Exporting all messages to a `.txt` file.

**Message Format Example:**

```
*GroupBuy Update 06/01 Order List*

#Pickup time: 3–7 PM
=====================
Buyer: Alice

Item: Milk Tea
Price: 35
Quantity: 2
---------------------
Item: Black Sugar Jelly
Price: 25
Quantity: 1
---------------------
Please confirm upon reading. Thank you!
```

**Usage:**

```bash
python 訊息.py
```

1. Click **Upload Excel File** and choose your `.xlsx` order sheet.
2. Browse the list of buyers.
3. Click to copy messages or edit them before export.
4. Use **Sort by Product** and **Search** to navigate.

---

### 3️⃣ `統計.py` – Comment Order Statistics Tool

**Description:**

A lightweight GUI tool to analyze comment-based orders using shorthand like `A+1`, `B 3`, or `C*1`. It extracts and sums product codes and quantities, then displays them in a table format.

**Key Features:**

* Recognizes patterns: `A+1`, `B 3`, `C*2`, etc.
* Aggregates total counts per product.
* Displays results in a scrollable table.
* One-click copy to clipboard in format like: `A.3,B.2,C.1`.
* Supports clearing all input and result fields.

**Usage:**

```bash
python 統計.py
```

1. Paste raw comment content in the input box.
2. Click **計算統計** ("Calculate Statistics") to process.
3. Click **複製結果** ("Copy Result") to export to clipboard.
4. Click **清除** ("Clear") to reset.

---

## 📂 File Naming Notes

While the script filenames use Traditional Chinese:

| Chinese Name | Suggested English Name |
| ------------ | ---------------------- |
| 表單.py        | `form_parser.py`       |
| 訊息.py        | `message_generator.py` |
| 統計.py        | `stats_counter.py`     |

You can rename them for convenience, but no functional change is needed.

---

## 📝 License

This project is provided for **personal, non-commercial use only**.
For redistribution or commercial licensing, please contact the original author.
