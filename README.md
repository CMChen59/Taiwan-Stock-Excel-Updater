# 台灣股票 Excel 收盤價自動更新工具

本專案提供一套高效且 100% 原生相容的自動化工具，用於自動抓取台灣證券交易所 (TWSE) 與櫃買中心 (OTC) 的股票每日收盤數據，並將數據精確寫入多個 Excel 試算表（`.xlsm` / `.xlsx`）中。

---

## 🌟 核心功能與特色

1. **超高速批次 API 查詢**
   - 透過證交所 MIS API（`getStockInfo.jsp`）一次性批次查詢所有試算表內的股票數據，僅需 ~0.1 秒。
   - 具備 Yahoo Finance API 備援機制，確保資料獲取穩定可靠。

2. **100% 完整保留 Excel 公式、圖案與 VBA 按鈕**
   - 使用 Windows 原生 `win32com.client` (Excel Native COM) 引擎處理寫入。
   - 避免 `openpyxl` 等第三方套件破壞含有 VBA 按鈕與繪圖物件（`drawing1.xml`）檔的相容性問題。
   - 完整保留 `B` 欄 (`=IF(...)`) 與 `H` 欄 (`=E5-J5`) 等動態計算公式與儲存格格式。

3. **區塊批量寫入與暫停自動計算**
   - 寫入時自動開啟 `xlCalculationManual` 暫停全檔公式連動計算，完成後再恢復自動計算。
   - 將 `I~M` 欄（成交張數、昨收、開盤、最高、最低）包成二維陣列一次性寫入，大幅降低 RPC 負擔，寫入速度提高 10 倍以上。

4. **智慧檔案過濾**
   - 自動過濾開啟中的 Excel 鎖定檔（`~$*.xlsm`）。
   - 自動過濾非股票明細檔（例如 `支出明細-2024.xlsx`）。
   - 支援單獨更新指定檔案或一次更新資料夾下所有股票試算表。

5. **綠色免安裝單一執行檔 (`.exe`)**
   - 支援經由 PyInstaller 打包為獨立執行檔，無須在目標電腦上安裝 Python 環境即可運行。

---

## 📊 更新儲存格對照表

| 欄位 | 數據內容 | 寫入方式 |
| :--- | :--- | :--- |
| **B3** | 自動更新時間戳記 | `※更新時間：YYYY/MM/DD HH:MM:SS` |
| **E 欄** | 當日成交價 | 數字 (float) |
| **I 欄** | 當日成交總張數 | 數字 (int, 單位: 張) |
| **J 欄** | 前一日收盤價 | 數字 (float) |
| **K 欄** | 當日開盤價 | 數字 (float) |
| **L 欄** | 當日最高價 | 數字 (float) |
| **M 欄** | 當日最低價 | 數字 (float) |

---

## 📂 專案檔案結構

```text
Stock/
├── update_stocks.py          # Python 自動更新主程式
├── 一鍵更新.bat               # Windows 批次執行腳本
├── 更新股票收盤價.exe         # 打包完成的綠色免安裝可執行檔
├── update_stocks.log         # 執行日誌紀錄檔
├── README.md                 # 專案說明文件
└── .gitignore                # Git 版本控制忽略檔
```

---

## 🚀 使用說明

### 方式一：直接雙擊執行檔 (`.exe`)
1. 將 `更新股票收盤價.exe` 放置於包含股票試算表（`.xlsm`）的資料夾中。
2. 雙擊執行 `更新股票收盤價.exe`。
3. 依選單輸入數字 (如 `0` 代表更新全部檔案)，或直接按 `Enter` 預設全部更新。

### 方式二：使用 Python 腳本
```bash
# 1. 安裝依賴套件
pip install pywin32 openpyxl requests pyinstaller

# 2. 執行更新程式
python update_stocks.py

# 3. 指定更新全部檔案 (非互動模式)
python update_stocks.py --all
```

### 方式三：重新打包 `.exe` 執行檔
```bash
python -m PyInstaller --noconfirm --onefile --console --name "更新股票收盤價" update_stocks.py
```

---

## 📝 系統需求
- **作業系統**：Windows 10 / 11
- **應用軟體**：已安裝 Microsoft Excel (需支援 COM 介面)
