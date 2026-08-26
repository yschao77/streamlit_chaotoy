"""CLI: 從 Drive 最新 Sitegiant UPC 檔填補空白 UPC，結果寫回同一資料夾。"""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from utils import apply_openpyxl_patch, run_sitegiant_upc_sync  # noqa: E402


def main():
    apply_openpyxl_patch()
    result = run_sitegiant_upc_sync()
    if not result.get("ok"):
        raise SystemExit(result.get("reason") or "UPC 填補失敗")
    if result.get("reason") == "no_updates":
        print(f"略過寫入：`{result.get('source')}` 沒有可填補的空白 UPC")
        return 0
    print(
        f"完成：來源 `{result.get('source')}` → `{result.get('output')}`，"
        f"填補 {result.get('updated')} / {result.get('source_rows')} 筆"
    )
    print("請將結果檔上傳回 Sitegiant 批量編輯 UPC。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
