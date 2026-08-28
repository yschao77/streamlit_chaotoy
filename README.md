# streamlit_chaotoy

麗嬰與蝦皮商務數據情報中心：Streamlit 進銷存工具。資料在 Google Drive 的 Excel / xlsm，沒有本機資料庫。

## 手冊

- [手冊總覽](docs/manuals/index.md)（[HTML](docs/manuals/index.html)）
- [使用者手冊](docs/manuals/user/index.md)（[HTML](docs/manuals/user/index.html)）
- [工程師手冊](docs/manuals/engineer/index.md)（[HTML](docs/manuals/engineer/index.html)）

HTML 可在檔案總管雙擊開啟，不必架網站。

## 本機啟動

```text
pip install -r requirements.txt
streamlit run app.py
```

憑證放在本機 `.streamlit/secrets.toml` 的 `textkey`（已列入 `.gitignore`），不要提交到 git。CLI / GitHub Actions 使用環境變數 `GOOGLE_SERVICE_ACCOUNT_JSON`。
