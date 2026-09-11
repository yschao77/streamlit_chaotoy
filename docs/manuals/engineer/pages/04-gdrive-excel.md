# Google Drive 與 Excel

## 本頁重點

- 來源目錄與選檔規則在 `utils.TRACKED_SOURCES`。
- 禁止 API 新建檔；沒有 `existing_file_id` 就停止。例外：Sitegiant_UPC 的 `batch_edit_upc_added_only.xlsx` 可用 `allow_create=True`。
- 寫回 xlsm：`keep_vba=True`。條碼先 `clean_barcode`，儲存格格式 `@`。

## 連線

- Service：`st.secrets["textkey"]` 或 `GOOGLE_SERVICE_ACCOUNT_JSON`
- 選最新檔：`pick_latest_gdrive_file`（解析檔名日期，不要只靠檔名排序）
- Sitegiant 來源可含 zip：`download_source_spreadsheet` 解出 xlsx
- 只讀可用 calamine；寫 xlsm 用 openpyxl + `keep_vba=True`

## TRACKED_SOURCES（用途）

| key | 資料夾意義 | 檔名慣例 | consumed |
|-----|------------|----------|----------|
| liying_master | 麗嬰採購產品總表 | `麗嬰採購產品總表.xlsm` | 是 |
| local_prod | 商品列表 | `商品列表.xlsx` | 是 |
| shopee_master | 蝦皮賣場商品列表 | `蝦皮賣場商品列表.xlsm` | 是 |
| price_summary | 價格統整表 | `商品蝦皮麗嬰價格統整表.xlsx` | 是 |
| sitegiant_upc | Sitegiant_UPC | `batch_edit_item_upc_assignment_all_DD-MM-YYYY-*.xlsx` | 是 |
| sitegiant_basic | Sitegiant_BasicInfo | `batch_edit_basic_info_all_DD-MM-YYYY-*.xlsx` | 是 |
| shopee_sales | 蝦皮_價格及庫存 | `mass_update_sales_info_3062950_YYYYMMDD*.xlsx` | 是 |
| shopee_media | 蝦皮_媒體資訊 | `mass_update_media_info_3062950_*.xlsx` | 否 |
| shopee_basic_info | 蝦皮_商品標題及描述 | `mass_update_basic_info_3062950_*.xlsx` | 否 |
| shopee_shipping | 蝦皮_配送選項 | `mass_update_shipping_info_3062950_*.xlsx` | 否 |
| shopee_unpublished | 蝦皮_未上架 | `mass_republish_items_3062950_*.xlsx` | 否 |
| preorder_orders | Sitegiant_Preorder_Orders | `Orders_DD-MM-YYYY-*.xlsx` | 是 |
| preorder_tracker | 預購追蹤.xlsx（file id） | `預購追蹤.xlsx` | 是 |
| sg_restock_template | Import Restock 官方殼（file id，只讀） | 後台下載的空殼 | 否 |

`consumed=False` 只出現在狀態頁，不寫入主表。`preorder_orders` 已改為 consumed：預購追蹤第2階看板會讀最新 `Orders_DD-MM-YYYY-*.xlsx`。本機上傳只預覽，禁止把 All Orders `files.create` 到 Drive。

常數定義在 `utils.py`（資料夾 ID 不是憑證）。另有 `ID_HISTORY_INWARD_FOLDER`、`ID_HISTORY_INWARD_INDEX`、`ID_DOWNLOAD_ROOT`、`ID_PRICE_SUMMARY_FALLBACK`、`ID_PREORDER_ORDERS_FOLDER`、`ID_PREORDER_TRACKER`、`ID_SG_RESTOCK_TEMPLATE`。

## 工作表

| 檔案 | 工作表 |
|------|--------|
| 麗嬰採購產品總表.xlsm | `麗嬰國際產品總表`、`已處理採購單`、`刪除紀錄`、`metadata` |
| 蝦皮賣場商品列表.xlsm | `蝦皮商品列表`、`匯入檔案` |
| 商品蝦皮麗嬰價格統整表.xlsx | `商品蝦皮麗嬰價格統整表` |
| 預購追蹤.xlsx | `活動`、`廠商單歷史` |
| 入庫明細索引.xlsx | `明細`、`來源檔` |

蝦皮校正來源為最新 mass_update；匯入去重用 md5。

歷史入庫：單據在 `YYYY/YYMM` 子資料夾（例如 `2026/2608`）。雲端沒有同名檔時只提供下載，不 `allow_create`。查詢用的 `入庫明細索引.xlsx` 放在歷史入庫**根目錄**，以 `ID_HISTORY_INWARD_INDEX` 覆寫；工作表為 `明細`、`來源檔`。

## 條碼

Excel 常把條碼變成 `471....0` 或科學記號。顯示與比對前一律 `clean_barcode`。
