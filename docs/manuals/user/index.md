# 使用者手冊總覽

本手冊對應畫面左側導覽：先選「核心管理模組」，再選「執行項目」。標題與 App 側邊欄相同。

## 本頁重點

- 資料都在 Google Drive，畫面只是讀寫那些 Excel。
- 建議每日先看雲端狀態，再校正蝦皮，再跑三表整合，最後做入庫或 UPC。
- 找不到檔案、重複匯入、憑證問題時，先對照該頁的「常見問題」。

## 系統是什麼

「麗嬰與蝦皮商務數據情報中心」幫你把三邊商品資料對在一起：

| 詞 | 意思 | 雲端檔 |
|----|------|--------|
| 麗嬰 | 採購／批發主檔 | `麗嬰採購產品總表.xlsm` |
| 蝦皮 | 賣場商品列表 | `蝦皮賣場商品列表.xlsm` |
| 內部商品列表 | 自家 SKU 清單 | `商品列表.xlsx`（資料夾「商品列表」） |
| 統整表 | 三表交叉後的價格與 SKU | `商品蝦皮麗嬰價格統整表.xlsx` |
| Sitegiant | 電商入庫／批量編輯 | `batch_edit_*.xlsx` 與入庫單 |

對鍵（查資料時常用）：

- 內部 `自定義編碼` = 蝦皮 `iSKU`
- 內部 `c` = 麗嬰 `條碼` = Sitegiant 國際條碼（UPC）
- 入庫單的「庫存貨品名稱」來自統整表 `sitegiant庫存SKU`，不是蝦皮商品名稱

## 如何打開

1. 用平常的網址打開 Streamlit 畫面（本機通常是 `http://localhost:8501`）。
2. 左側選模組，再選子頁。
3. 頁面上方的「目前導覽路徑」可確認你在哪一頁。

若整頁出現「憑證缺失／無效」或無法列出雲端檔，請找工程師檢查 Secrets 與 Drive 權限，不要自行改檔名亂試。

## 建議每日順序

1. **各表最新修改時間** — 確認今日來源檔已上傳。
2. **蝦皮商品清單轉換** — 校正最新 mass_update（台北 10:00 後若尚未匯入會自動跑一次）。
3. **PowerQuery 執行三表整合** — 重算統整表並回寫雲端（入庫會用到這張表）。
4. 依需要：**Sitegiant 採購入庫單格式轉換** 或 **批量新增 UPC**。
5. 有新麗嬰採購單時：先 **麗嬰商品表合併和與審核**，再跑三表整合。

```
雲端狀態 → 蝦皮 iSKU 校正 → 三表整合 → 入庫轉換
                ↘ 批量 UPC
```

## 依畫面標題

### 雲端資料狀態

- [各表最新修改時間](pages/01-cloud-status.md)

### Sitegiant 電商整合管理

- [Sitegiant 採購入庫單格式轉換](pages/02-sitegiant-inward.md)
- [Sitegiant 歷史入庫單紀錄](pages/03-sitegiant-history.md)
- [查詢入庫紀錄](pages/12-inward-query.md)
- [Sitegiant 批量新增UPC](pages/04-sitegiant-upc.md)
- [採購單待處理](pages/05-pending-po.md)
- [預購追蹤](pages/13-preorder-tracker.md)

### 商品蝦皮麗嬰統整管理

- [PowerQuery 執行三表整合](pages/06-powerquery.md)
- [麗嬰商品表合併和與審核](pages/07-liying-merge.md)
- [蝦皮商品清單轉換](pages/08-shopee-convert.md)
- [PowerQuery 三表整合歷史紀錄](pages/09-powerquery-history.md)
- [商品清單紀錄查詢](pages/10-product-query.md)
- [麗嬰商品總表數據查詢](pages/11-liying-query.md)
