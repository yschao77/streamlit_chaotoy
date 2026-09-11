# 架構與資料流

## 本頁重點

- Streamlit 單頁應用；側邊欄 `selectbox` 選模組、`radio` 選子頁。
- Drive／Excel 的唯一實作在 `utils.py`；`views/` 只 render。
- 三表整合與蝦皮／UPC 同步，畫面與 CLI 走同一組函式。

## 執行時資料流

1. `app.py` 設定 `page_config`、套用 openpyxl patch、用關鍵字查出主檔 ID。
2. 組成 `gdrive_cfg`（或個別 Drive ID）傳給 `views/*.render`。
3. 子頁開頭固定：`st.title` + 麵包屑，不再做第二套導覽。
4. 讀檔：`download_gdrive_file_to_bytes` / `get_cached_gdrive_file_bytes`；寫檔：`upload_or_update_gdrive_file`。

```
使用者
  → app.py 路由
    → views/status_page.py | sitegiant_page.py | integration_page.py
      → utils.py（Drive API、清洗、merge、sync）
        → Google Drive Excel/xlsm
```

## 三個模組

| 側邊欄模組 | 檔案 | 子頁 |
|------------|------|------|
| 雲端資料狀態 | `views/status_page.py` | 各表最新修改時間 |
| Sitegiant 電商整合管理 | `views/sitegiant_page.py` | 入庫轉換、歷史入庫、查詢入庫、批量 UPC、待處理、預購追蹤 |
| 商品蝦皮麗嬰統整管理 | `views/integration_page.py` | 三表整合、麗嬰合併審核、蝦皮轉換、歷史／查詢 |

新子頁：先在 `app.py` 的 `radio` 加字串，再在對應 `views/*.py` 加 `if sub_page == "..."`。字串必須完全一致。

`status_page.render(sub_page, gdrive_cfg)`、`integration_page.render(sub_page, gdrive_cfg)`；`sitegiant_page` 傳個別 ID。

## 對鍵與計算

`build_price_summary_df`：

- 商品列表 ⟕ 蝦皮（`自定義編碼` = `iSKU`）
- 再 ⟕ 麗嬰（`c` = `條碼`）
- `sitegiant庫存SKU`：`c` 對 batch_edit `國際條碼（UPC）` 的「庫存貨品名稱」

入庫單「庫存貨品名稱」只抓 `sitegiant庫存SKU`，不要用蝦皮商品名稱。

財務欄：零售八折 `× 0.8`、未稅 `÷ 1.05`、稅款 = 含稅 − 未稅。

## CLI 與畫面共用

| 畫面動作 | utils 函式 | script |
|----------|------------|--------|
| 蝦皮 iSKU 校正 | `run_shopee_isku_sync` / `apply_shopee_isku_from_source` | `scripts/sync_shopee_from_gdrive.py` |
| 三表整合回寫 | `run_price_summary_sync` / `build_price_summary_df` | `scripts/sync_price_summary.py` |
| 填 UPC | `run_sitegiant_upc_sync` / `fill_sitegiant_upc` | `scripts/sync_sitegiant_upc.py` |
| 狀態表 | `collect_tracked_file_status` | `scripts/report_gdrive_status.py` |

`temp/` 是舊版 `app_*.py` 快照，只讀，不當實作來源。
