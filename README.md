# 團購管理工具套件（Group Buying Toolkit）

這是一個桌面應用程式工具包，專為團購主設計，用於簡化團購過程中從留言訂單整理、個人通知訊息產生到數量統計等繁瑣工作。

> 📌 僅供個人使用。若欲商務或再發佈，請與作者聯繫授權。


## 📦 專案特色功能

本專案包含三個獨立工具模組：

### 1. 📋 團購訂單表產生器（`表單.py`）

- 從社群貼文留言中自動解析每位買家的訂購項目與數量。
- 自動整理成買家 × 商品的交叉表格式。
- 支援重複留言合併、異常訊息過濾。
- 可匯出為 Excel 檔案供後續處理。

### 2. ✉️ 訂購訊息生成器（`訊息.py`）

- 從彙總後的 Excel 訂單中，自動生成每位買家的訂單通知訊息。
- 支援搜尋訂購人、依產品排序買家名單。
- 可一鍵複製訊息、進行個別編輯、批量匯出訊息為文字檔。

### 3. 🧮 訂單統計工具（`統計.py`）

- 快速統計貼文留言中各商品（如 A+1、B*2）總數。
- 提供「依商品代碼」與「總訂單量」兩種統計模式。
- 支援一鍵複製統計結果。

---

## 💻 安裝與執行

### 1. 安裝環境需求

- Python 3.8 或以上
- 系統需支援 Tkinter GUI（macOS 請先安裝 XQuartz）

### 2. 安裝必要套件

在終端機中執行以下指令：

```bash
pip install -r requirements.txt
````

### 3. 啟動工具

可根據需要個別執行下列腳本：

```bash
python 表單.py     # 啟動訂單表產生器
python 訊息.py     # 啟動訂購訊息生成器
python 統計.py     # 啟動留言統計工具
```

---

## 📁 檔案說明

| 檔案名稱             | 功能簡介                                            |
| ---------------- | ----------------------------------------------- |
| 表單.py            | 建立團購訂單彙總表並匯出 Excel                              |
| 訊息.py            | 根據 Excel 自動產生買家通知訊息                             |
| 統計.py            | 對留言中商品數量進行統計                                    |
| requirements.txt | 套件安裝清單（pandas、openpyxl、customtkinter、pyperclip） |

---

## 📝 使用注意事項

* 若需顯示中文字建議修改預設字型，可在程式開頭將 `Segoe UI` 改為 `Microsoft JhengHei` 或其他中文字型。
* 訂單表中的商品價格欄需手動填寫，否則訊息將顯示「無價格」。
* 所有工具皆為單機執行，無需網路連線或伺服器環境。

---

## 👤 作者資訊

* GitHub: [August980422](https://github.com/August980422)
* 專案開發與設計：August（個人製作）
* 使用問題與授權洽詢請透過 GitHub 聯絡

---

## 📄 授權條款

本工具僅供個人非商業用途使用。
禁止未經授權進行以下行為：

* 重新上傳本專案至其他平台
* 用於任何營利用途
* 修改後公開發布未註明原始出處

如需進行以上用途，請務必與原作者聯繫取得授權。

---

感謝使用，祝您團購順利！

```
