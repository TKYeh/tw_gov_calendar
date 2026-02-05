TW_gov_calendar
=================

此專案用來從 data.gov.tw 自動下載「中華民國政府行政機關辦公日曆」CSV，並轉換成 iCalendar (.ics) 檔案。

***
訂閱最新行事曆網址為: https://raw.githubusercontent.com/TKYeh/tw_gov_calendar/refs/heads/main/TW_gov_calendar.ics
***

主要功能
- 下載 dataset 中各年度的 CSV（自動偵測編碼，支援 UTF-8 / Big5 等）
- 正規化欄位名稱（支援 `西元日期`、`假日類別`、`備註` 及常見替代欄名）
- 過濾：移除「例假日」或只保留有備註的紀錄（遇不同 schema 時會自動處理）
- 產生每年度的 `.ics`（輸出到 `output/`）
- 自動將對應「今年」的 `.ics` 複製為 `TW_gov_calendar.ics`（放在工作目錄）

使用方法
1. 本地執行：
   ```bash
   python generate_calendar.py
   ```
   執行完成後會在 `output/` 看到每年度的 `.ics`，並在專案根目錄得到 `TW_gov_calendar.ics`（對應今年）。

2. 自動化（GitHub Actions）：
   - 已包含工作流程 `.github/workflows/scheduled_update.yml`，每天會檢查一次，僅在台北時間每月第一天或最後一天執行更新，並自動 commit 輸出檔案。

注意事項
- 若 dataset CSV 欄位命名異常（例如編碼錯誤），程式會嘗試偵測編碼並正規化欄位，若仍無法對應欄位會跳過該檔案並輸出警告。
- 若要修改選取哪個年度為 `TW_gov_calendar.ics`，請參考 `get_ics_for_current_year` 中的邏輯，或在 `main()` 修改選擇條件。

Requirements
--
安裝程式所需的 Python 套件：包含 `pandas`, `requests`, `ics`, `chardet`

```bash
python -m pip install -r requirements.txt
```
