# 工程師手冊總覽

本手冊說明如何啟動、改程式、接 Drive，以及定時同步。資料在 Google Drive 的 Excel／xlsm，沒有本機資料庫。這是公開 GitHub repo：憑證與 private_key 不得進 git、issue 或畫面。

## 本頁重點

- `app.py` 只負責 page_config 與側邊欄路由；讀寫實作在 `utils.py`。
- Streamlit 用 `st.secrets["textkey"]`；CLI／Actions 用 `GOOGLE_SERVICE_ACCOUNT_JSON`。
- 禁止 API 新建雲端檔（例外：`batch_edit_upc_added_only.xlsx`）。

## 架構

```
app.py          側邊欄：模組 → 子頁，組 gdrive_cfg
views/          畫面 render；從 utils 呼叫
utils.py        Drive I/O、TRACKED_SOURCES、條碼、主表／蝦皮寫入
scripts/        與畫面同一套同步函式，給 CLI 與 GitHub Actions
```

對鍵：

- 內部 `自定義編碼` ↔ 蝦皮 `iSKU`
- 內部 `c` ↔ 麗嬰 `條碼` ↔ batch_edit `國際條碼（UPC）` → 統整表 `sitegiant庫存SKU`

商業數字未要求不要改：稅率 `1.05`、零售八折 `0.8`、UID 後備 `3473`。

## 本機最短路徑

```text
pip install -r requirements.txt
streamlit run app.py
```

憑證：本機 `.streamlit/secrets.toml` 的 `textkey`（已 gitignore）。每個 Drive 資料夾要把 service account 的 `client_email` 加成檢視者或編輯者。

## 依主題

- [架構與資料流](pages/01-architecture.md)
- [本機啟動與憑證](pages/02-setup.md)
- [檔案職責](pages/03-file-roles.md)
- [Google Drive 與 Excel](pages/04-gdrive-excel.md)
- [定時同步](pages/05-ci-sync.md)
- [開發慣例與故障排除](pages/06-conventions-troubleshooting.md)

使用者畫面步驟見 [使用者手冊](../user/index.md)。
