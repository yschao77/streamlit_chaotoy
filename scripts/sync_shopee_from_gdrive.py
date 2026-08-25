"""CLI: 從 Drive 抓最新 mass_update 並校正蝦皮賣場商品列表.xlsm。"""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from utils import (  # noqa: E402
    apply_openpyxl_patch,
    apply_shopee_isku_from_source,
    download_gdrive_file_to_bytes,
    list_gdrive_files,
    pick_latest_gdrive_file,
)

ID_SHOPEE_FOLDER = os.environ.get("ID_SHOPEE_FOLDER", "17eiGnXyU4KwNS6IR5bubBPti46SKXMH0")
ID_SHOPEE_MASS_UPDATE_FOLDER = os.environ.get(
    "ID_SHOPEE_MASS_UPDATE_FOLDER", "1OG2kpiLmhBjMR-12vDPUkEdGbUKs4vLs"
)


def _find_shopee_master():
    files = list_gdrive_files(ID_SHOPEE_FOLDER, name_contains="蝦皮賣場商品列表")
    if not files:
        return None, None
    files.sort(key=lambda x: x.get("modifiedTime") or "", reverse=True)
    return files[0]["id"], files[0]["name"]


def main():
    apply_openpyxl_patch()
    latest = pick_latest_gdrive_file(
        ID_SHOPEE_MASS_UPDATE_FOLDER, "mass_update_sales_info_3062950", "mass_update"
    )
    if not latest:
        raise SystemExit("找不到 mass_update_sales_info_3062950_YYYYMMDD*.xlsx")

    master_id, master_name = _find_shopee_master()
    if not master_id:
        raise SystemExit("找不到雲端檔案：蝦皮賣場商品列表.xlsm")

    print(f"來源: {latest['name']}")
    print(f"主表: {master_name} ({master_id})")
    file_bytes = download_gdrive_file_to_bytes(latest["id"])
    result = apply_shopee_isku_from_source(
        file_bytes, latest["name"], master_id, ID_SHOPEE_FOLDER, master_name
    )
    if result["reason"] == "duplicate":
        print(f"略過：{result['name']} 已匯入（md5 重複）")
        return 0
    if not result["ok"]:
        raise SystemExit(f"校正失敗: {result['reason']}")
    print(f"完成：{result['name']} @ {result['imported_at']}，共 {len(result['df'])} 筆")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
