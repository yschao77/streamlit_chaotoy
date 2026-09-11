# 麗嬰與蝦皮商務數據情報中心 — 手冊總覽

這套 Streamlit 工具用來處理進銷存資料。真正的資料在 **Google Drive 的 Excel / xlsm**，沒有本機資料庫。

請依角色打開對應手冊。每一本都有「總覽」和「依畫面標題分頁」；HTML 可直接用瀏覽器開啟，不必架網站。

## 本頁重點

- 使用者：依側邊欄標題操作各頁。
- 工程師：啟動、憑證、Drive 讀寫與定時同步。
- 兩種格式：Markdown（GitHub）與 HTML（本機瀏覽器）。

## 選擇手冊

| 手冊 | Markdown | HTML | 給誰看 |
|------|----------|------|--------|
| [使用者手冊](user/index.md) | [user/index.md](user/index.md) | [user/index.html](user/index.html) | 日常操作、匯入、查詢、入庫 |
| [工程師手冊](engineer/index.md) | [engineer/index.md](engineer/index.md) | [engineer/index.html](engineer/index.html) | 啟動、架構、CI、故障排除 |

## 系統是什麼

側邊欄有三個核心模組：

1. **雲端資料狀態**：看各 Drive 來源最新檔時間。
2. **Sitegiant 電商整合管理**：採購入庫轉換、歷史入庫、查詢入庫、批量 UPC、待處理清單、預購追蹤。
3. **商品蝦皮麗嬰統整管理**：麗嬰採購單合併審核、蝦皮 iSKU 校正、三表整合與查詢。

常用詞：

- **麗嬰**：採購／批發主檔（`麗嬰採購產品總表.xlsm`）。
- **蝦皮**：賣場商品列表（`蝦皮賣場商品列表.xlsm`）。
- **Sitegiant**：電商入庫與批量編輯檔格式。
- **三表整合**：內部商品列表 × 蝦皮 × 麗嬰 → `商品蝦皮麗嬰價格統整表`。

## 如何打開 HTML

在檔案總管進入 `docs/manuals/`，雙擊 `index.html`。相對連結即可在離線環境翻頁。
