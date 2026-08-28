# Sitegiant 批量新增UPC

路徑：Sitegiant 電商整合管理 → Sitegiant 批量新增UPC

## 本頁重點

- 用已校正的蝦皮列表（iSKU → GTIN）填 Sitegiant 空白 UPC。
- 只填「目前空白」的列，不會改已有 UPC。
- 結果下載後，再上傳回 Sitegiant 批量編輯。

## 何時使用

Sitegiant 商品缺國際條碼，而蝦皮列表已有對應 GTIN。

## 前置條件

1. 已確認 iSKU 對應 UPC／GTIN。
2. 把 Sitegiant 批量編輯 UPC 檔放到雲端 `Sitegiant_UPC`（xlsx 或 zip）：
   - `batch_edit_item_upc_assignment_all_DD-MM-YYYY-*.xlsx`
3. 已在「蝦皮商品清單轉換」校正並回寫蝦皮主表。

## 操作步驟

1. 確認頁面顯示最新 UPC 來源檔名稱與時間。
2. 按 **從雲端最新 UPC 檔填補**。
3. 看成功筆數與預覽。
4. 下載結果，或到雲端取 `batch_edit_upc_added_only.xlsx`。
5. 上傳回 Sitegiant 覆蓋。

備用：展開「手動上傳 UPC 檔」，上傳本機 xlsx／xls／csv／zip 再執行填補（手動上傳預設不寫回雲端結果檔）。

## 預期檔案

- 讀：最新 `batch_edit_item_upc_assignment_all_*.xlsx`（或 zip 內同名 xlsx）
- 讀：`蝦皮賣場商品列表.xlsm` 的「蝦皮商品列表」
- 寫：`batch_edit_upc_added_only.xlsx`（定時同步也會寫這份）

寫回雲端時，service account 需要對 `Sitegiant_UPC` 資料夾有編輯者權限。

## 常見問題

- **無法讀取蝦皮商品列表**：先做蝦皮商品清單轉換。
- **找不到 batch_edit_item_upc_assignment_all_…**：檔名或日期格式不對，或資料夾是空的。
- **未找到任何可填補的缺失 UPC**：空白列的 SKU 在蝦皮對不到有效 GTIN，或本來就沒有空白 UPC。
- **找不到 SKU 或 UPC 欄**：匯出檔欄名不是 Item SKU / UPC 等系統認得的名稱。
