# 開發慣例與故障排除

## 本頁重點

- 共用邏輯只改 `utils.py`；頁面不要重寫 `clean_barcode`。
- Drive service：`@st.cache_resource`；檔案 bytes／主表：`@st.cache_data(ttl=600)`（ID 查詢可用 3600）。寫入成功後要清相關 cache。
- 使用者可見的憑證錯誤維持「憑證缺失／無效」，不要 `st.error(str(e))` 若可能含金鑰。

## 慣例

```python
# 好
from utils import clean_barcode
df["條碼"] = df["條碼"].map(clean_barcode)

# 不好：在頁面再定義一份清洗
```

- 使用者看得到的失敗用 `st.error`（可加 `st.info`）；不要空的 `except: pass`
- Drive 連線失敗可 `st.stop()`
- `views/` 既有的 `sys.path` 補丁可保留，新檔不要再複製一份

## 快取

寫雲端成功後常見要清：

- `load_master_data.clear()`
- `load_shopee_data.clear()`
- `get_cached_gdrive_id.clear()`
- `get_cached_gdrive_file_bytes.clear()`
- 狀態頁：`_cached_status.clear()`

否則畫面仍是舊檔。

## 故障排除

| 現象 | 先查 |
|------|------|
| 憑證缺失／無效 | Secrets `textkey` 或 Actions Secret 是否存在且 JSON 有效；private_key 換行是否被弄壞（utils 會嘗試修復 PEM） |
| `files.list` 空 | service account 沒被加進該資料夾 |
| 找不到 mass_update / batch_edit | 檔名日期格式；`pick_latest_gdrive_file` 的 kind（ymd / dmy） |
| 統整表無法回寫 | 雲端沒有既有主檔；API 不新建 |
| 歷史入庫只下載不覆寫 | 資料夾沒有同名檔 |
| UPC 填補 0 筆 | 蝦皮列表空或沒有效 GTIN；SKU 對不上 |
| 重複匯入被拒 | md5 已在「匯入檔案」或「已處理採購單」 |
| 條碼對不上 | 未經 `clean_barcode`；Excel 科學記號 |
| 畫面資料過期 | 寫入後沒清 cache，或 TTL 未到 |

監看用的蝦皮資料夾沒接流程是正常的，不要在同步 script 裡誤讀它們當主表來源。
