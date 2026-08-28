# 定時同步

## 本頁重點

- Workflow：`.github/workflows/sync-shopee.yml`
- 排程：cron `0 2 * * *`（UTC 02:00，約台北 10:00），也可 `workflow_dispatch`
- 憑證：repo Secret `GOOGLE_SERVICE_ACCOUNT_JSON`，用環境變數注入，不要把 JSON 寫死在 YAML

## 步驟順序

前三步 `continue-on-error: true`；最後狀態報告 `if: always()`。

1. `python scripts/sync_shopee_from_gdrive.py`  
   最新 mass_update → 校正 iSKU → 覆寫蝦皮主表。md5 重複則略過。
2. `python scripts/sync_price_summary.py`  
   三表 + BasicInfo → 覆寫價格統整表。雲端必須已有統整表主檔。
3. `python scripts/sync_sitegiant_upc.py`  
   最新 UPC 檔填空白 → 寫 `batch_edit_upc_added_only.xlsx`。沒有可填則略過寫入。
4. `python scripts/report_gdrive_status.py`  
   列出各來源最新檔與台北時間；寫入 GitHub Step Summary。

本機手動跑同一支 script 即可，需先設好 `GOOGLE_SERVICE_ACCOUNT_JSON`。

## 可選郵件

`report_gdrive_status.py` 在同時有 `STATUS_MAIL_TO` 與 `SMTP_HOST` 才寄信。可選：`SMTP_PORT`（預設 587）、`SMTP_USER`、`SMTP_PASSWORD`、`SMTP_FROM`。沒有這些 Secret 時不會寄信，也不要把帳密寫進程式。

畫面「各表最新修改時間」的郵件區塊標示尚未啟用；狀態仍會進 Actions 摘要。

## 失敗時

個別 sync 失敗不會擋住後面的狀態報告。到 Actions log 看該步的 `raise SystemExit(reason)` 訊息（找不到檔、校正失敗等），不要把憑證內容貼進 issue。
