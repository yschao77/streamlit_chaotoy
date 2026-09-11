# 檔案職責

## 本頁重點

- 新邏輯放 `utils.py` 或 `views/`，不要再把整份 app 複製進 `temp/`。
- 共用函式只維護一份：頁面 `from utils import ...`，不要在 `app.py` 再定義同名函式。
- `save_to_master_xlsm` / `save_to_shopee_master_xlsm` 以 `utils.py` 簽名為準。

## 對照

| 路徑 | 職責 |
|------|------|
| `app.py` | `page_config`、側邊欄路由、組 `gdrive_cfg` 傳子頁 |
| `utils.py` | Drive I/O、`TRACKED_SOURCES`、條碼清洗、主表／蝦皮 xlsm 讀寫、三表／iSKU／UPC 同步 |
| `views/status_page.py` | 各雲端來源最新修改時間 |
| `views/sitegiant_page.py` | 入庫轉換、歷史入庫、查詢入庫、批量 UPC、待處理、預購追蹤 |
| `views/integration_page.py` | 三表整合、麗嬰合併審核、蝦皮轉換、查詢 |
| `scripts/` | 薄包裝，呼叫 `utils` 的 `run_*_sync` |
| `.github/workflows/sync-shopee.yml` | 每日排程跑 scripts |
| `temp/` | 舊快照，只讀 |
| `GDriveFolder_ID.md` | 資料夾 ID 與檔名慣例備忘 |

注意：`app.py` 仍有一份舊的 `save_to_master_xlsm`、`clean_barcode` 等定義；**頁面實際 import 的是 utils**。改行為時改 `utils.py`，不要只改 `app.py` 裡重複的函式。

## UI 慣例

- `st.set_page_config(..., layout="wide")` 只在 `app.py`
- UI 繁中 + 既有 emoji 標籤
- 主操作按鈕 `use_container_width=True`
- 跨 rerun 資料放 `st.session_state`（例如 `inward_input_df`、`pq_result`），key 保持穩定

## 不要改的數字（未要求時）

- 稅率 `1.05`
- 零售八折 `0.8`
- UID 後備 `3473`（metadata 空時）
