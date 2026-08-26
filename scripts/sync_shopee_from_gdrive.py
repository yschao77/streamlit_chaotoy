"""CLI: 從 Drive 抓最新 mass_update 並校正蝦皮賣場商品列表.xlsm。"""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from utils import apply_openpyxl_patch, run_shopee_isku_sync  # noqa: E402


def main():
    apply_openpyxl_patch()
    result = run_shopee_isku_sync()
    if result.get("reason") == "duplicate":
        print(f"略過：{result.get('name')} 已匯入（md5 重複）")
        return 0
    if not result.get("ok"):
        raise SystemExit(result.get("reason") or "校正失敗")
    print(f"來源: {result.get('source')}")
    print(f"主表: {result.get('master')}")
    print(f"完成：{result['name']} @ {result['imported_at']}，共 {len(result['df'])} 筆")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
