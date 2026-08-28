# 本機啟動與憑證

## 本頁重點

- Python 3.11、`requirements.txt`、`streamlit run app.py`。
- 本機 Secrets：`st.secrets["textkey"]`；不要印出、不要 commit。
- CLI／Actions：環境變數 `GOOGLE_SERVICE_ACCOUNT_JSON`（整份 JSON 字串）。

## 啟動

```text
pip install -r requirements.txt
streamlit run app.py
```

Dev Container（`.devcontainer/devcontainer.json`）會安裝依賴並在 8501 跑 Streamlit。`packages.txt` 目前只有 `python3-tk`。

依賴：`streamlit`、`pandas`、`openpyxl`、Google Drive API 套件、`python-calamine`、`xlrd`。不要無故加套件。

## 憑證放哪裡

| 環境 | 來源 |
|------|------|
| Streamlit 本機 | `.streamlit/secrets.toml` 的 `textkey`（gitignore） |
| CLI / GitHub Actions | `GOOGLE_SERVICE_ACCOUNT_JSON` |
| 備援 | `GOOGLE_APPLICATION_CREDENTIALS` 指向本機 JSON 路徑 |

`load_google_service_account_info()` 依上述順序讀取，只在記憶體使用。

**禁止：**

- 把 service account JSON 寫進 workflow YAML、README、手冊範例的真實內容
- `print` 環境變數或 `st.write(st.secrets["textkey"])`
- `git add -f` secrets、`*.pem`、`*service-account*.json`
- 錯誤訊息洩漏金鑰；使用者可見字串用「憑證缺失／無效」

Drive **資料夾 ID** 可以寫在 `utils.py`；**private_key 不行**。

## Drive 權限

新資料夾必須把 service account 的 `client_email` 加成檢視者，否則 `files.list` 會是空的。

需要寫回的資料夾（統整表、麗嬰／蝦皮主表、歷史入庫、Sitegiant_UPC）要編輯者。

## 本機不要提交的匯出檔

`.gitignore` 已排除例如 `batch_edit_basic_info_all_*.xlsx`、`mass_update_sales_info_*.xlsx`。不要 `git add -f` 這些生產匯出。
