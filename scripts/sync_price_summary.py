"""CLI: 從 Drive 最新三表 + Sitegiant BasicInfo 重算並覆寫價格統整表。"""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from utils import apply_openpyxl_patch, run_price_summary_sync  # noqa: E402


def main():
    apply_openpyxl_patch()
    result = run_price_summary_sync()
    if not result.get("ok"):
        raise SystemExit(result.get("reason") or "三表整合失敗")
    print(
        f"完成：{result.get('summary')}，共 {result.get('rows')} 筆"
        + (f"｜batch_edit `{result.get('batch')}`" if result.get("batch") else "")
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
