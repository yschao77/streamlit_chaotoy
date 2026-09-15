import sys
import os

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import streamlit as st
import pandas as pd
import io
import datetime
import re

# 從外層的 utils.py 引入我們需要的工具
from utils import (
    HAS_CALAMINE, 
    download_gdrive_file_to_bytes, 
    upload_or_update_gdrive_file,
    get_cached_gdrive_file_bytes,
    clean_barcode,
    load_shopee_data,
    load_barcode_to_sitegiant_name_map,
    apply_sitegiant_name_from_summary,
    pick_latest_gdrive_file,
    download_source_spreadsheet,
    read_tabular_file,
    fill_sitegiant_upc,
    format_gdrive_time,
    extract_xlsx_from_zip,
    resolve_named_file,
    UPC_FILLED_FILENAME,
    HISTORY_INWARD_INDEX_NAME,
    list_history_inward_files,
    history_inward_file_label,
    history_inward_option_map,
    match_history_inward_file,
    load_history_inward_index,
    refresh_history_inward_index,
    invalidate_history_inward_caches,
    load_preorder_tracker,
    save_preorder_tracker,
    probe_sg_restock_template,
    preorder_tracker_bytes,
    PREORDER_TRACKER_NAME,
    load_preorder_orders,
    build_preorder_board,
    build_preorder_arrival_copy,
    _preorder_text,
    parse_vendor_po_bytes,
    upsert_vendor_rows_into_campaign,
    duplicate_preorder_skus,
    preorder_rows_sku_equals_item_no,
    lock_preorder_campaign,
    append_preorder_vendor_history,
    fill_vendor_po_qty_bytes,
    fill_sg_restock_bytes,
    taipei_now,
    PREORDER_VENDOR_HISTORY_COLUMNS,
    PREORDER_SKU_STATUS_COLUMNS,
    PREORDER_STATUS_PREORDER,
    PREORDER_STATUS_ARRIVED,
    extract_preorder_pending_df,
    attach_preorder_notify_days,
    upsert_preorder_sku_status,
    stamp_preorder_first_notify,
    preorder_notify_keys_from_lines,
    campaign_df_from_arrived_skus,
    arrived_skus_missing_from_campaign,
    ensure_preorder_sku_status_df,
)

def _inward_excel_bytes(df, sheet_name="SiteGiant入庫單"):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buf.getvalue()


PREORDER_CAMPAIGN_UI_COLUMN_ORDER = (
    "月份",
    "結單日",
    "SKU",
    "自購",
    "上限",
    "品名",
    "條碼",
    "貨號",
    "廠商檔名",
    "私密連結",
    "預計關閉",
    "實際關閉",
)
PREORDER_LOCK_DISPLAY_COLUMNS = ("SKU", "條碼", "貨號", "自購", "客戶量", "上限", "鎖定數量")


def _clear_preorder_campaign_editor():
    st.session_state.pop("preorder_campaign_editor_v4", None)
    st.session_state.pop("preorder_campaign_editor_v3", None)
    st.session_state.pop("preorder_campaign_editor_v2", None)
    st.session_state.pop("preorder_campaign_editor", None)


def _empty_preorder_sku_labels(campaign_df):
    labels = []
    if campaign_df is None or getattr(campaign_df, "empty", True):
        return labels
    for _, camp in campaign_df.iterrows():
        if _preorder_text(camp.get("SKU")):
            continue
        name = _preorder_text(camp.get("品名"))
        item_no = _preorder_text(camp.get("貨號"))
        barcode = clean_barcode(camp.get("條碼"))
        if not name and not item_no and not barcode:
            continue
        label = name or "（無品名）"
        if item_no:
            label += f"／貨號 `{item_no}`"
        elif barcode:
            label += f"／條碼 `{barcode}`"
        labels.append(label)
    return labels


def _save_preorder_tracker_button(campaign_df, key, *, primary=True):
    clicked = st.button(
        "覆寫雲端預購追蹤",
        type="primary" if primary else "secondary",
        use_container_width=True,
        key=key,
    )
    if not clicked:
        return
    result = save_preorder_tracker(
        campaign_df,
        st.session_state.get("preorder_vendor_history_df"),
        st.session_state.get("preorder_other_sheets"),
        file_name=st.session_state.get("preorder_tracker_name"),
        pending_df=st.session_state.get("preorder_pending_df"),
        sku_status_df=st.session_state.get("preorder_sku_status_df"),
        notify_df=st.session_state.get("preorder_notify_df"),
    )
    if result.get("ok"):
        st.session_state["preorder_save_ok"] = True
        get_cached_gdrive_file_bytes.clear()
        st.session_state.pop("preorder_loaded", None)
        st.rerun()
    st.error(f"❌ 寫回失敗：{result.get('reason') or '未知錯誤'}")


def _show_preorder_lock_tables(result):
    skipped = result.get("skipped") or []
    locked_rows = result.get("locked_rows") or []
    if skipped:
        st.warning(f"{len(skipped)} 列因未填 SKU 未鎖定。")
        st.dataframe(pd.DataFrame(skipped), hide_index=True, use_container_width=True)
    if locked_rows:
        show = pd.DataFrame(locked_rows)
        cols = [c for c in PREORDER_LOCK_DISPLAY_COLUMNS if c in show.columns]
        st.dataframe(show[cols], hide_index=True, use_container_width=True)


def _apply_preorder_lock_to_session(result):
    st.session_state["preorder_lock_result"] = result
    st.session_state["preorder_campaign_df"] = result["campaign"]
    if "preorder_vendor_history_before_lock" not in st.session_state:
        st.session_state["preorder_vendor_history_before_lock"] = st.session_state.get(
            "preorder_vendor_history_df"
        )
    st.session_state["preorder_vendor_history_df"] = append_preorder_vendor_history(
        st.session_state.get("preorder_vendor_history_before_lock"),
        result.get("history_add"),
    )
    _clear_preorder_campaign_editor()
    st.session_state.pop("preorder_vendor_fill", None)
    st.session_state.pop("preorder_restock_fill", None)
    st.session_state.pop("preorder_lock_preview", None)


def _render_preorder_orders_board(campaign_df):
    st.markdown("### 3. 對客戶訂單")
    st.caption("摘要看板與 SKU 狀態。本機 Orders 只預覽、不會上傳到 Drive。")

    reload_orders = st.button(
        "🔄 重新載入雲端 All Orders",
        use_container_width=True,
        key="preorder_orders_reload",
    )
    if reload_orders:
        get_cached_gdrive_file_bytes.clear()
        st.session_state.pop("preorder_orders_loaded", None)
        st.session_state.pop("preorder_orders_local_cache", None)
        st.session_state.pop("preorder_orders_local_key", None)
        st.session_state.pop("preorder_pending_source", None)
        st.rerun()

    if "preorder_orders_loaded" not in st.session_state:
        with st.spinner("⏳ 正在載入雲端最新 All Orders…"):
            st.session_state["preorder_orders_loaded"] = load_preorder_orders()
    drive_loaded = st.session_state.get("preorder_orders_loaded") or {}
    drive = drive_loaded.get("drive") or {}

    has_local = bool(st.session_state.get("preorder_orders_local_cache"))
    uploaded = None
    with st.expander("本機上傳 All Orders 預覽（不會寫入 Drive）", expanded=has_local):
        uploaded = st.file_uploader(
            "本機上傳 All Orders 預覽（不會上傳到 Drive）",
            type=["xlsx", "xls"],
            key="preorder_orders_local",
            help="覆蓋本次畫面使用的 Orders；不會 files.create／update 到雲端。",
        )
    if uploaded is not None:
        local_key = (uploaded.name, uploaded.size)
        if st.session_state.get("preorder_orders_local_key") != local_key:
            st.session_state["preorder_orders_local_cache"] = load_preorder_orders(
                uploaded.getvalue(),
                uploaded.name,
                drive=drive,
            )
            st.session_state["preorder_orders_local_key"] = local_key
        loaded_orders = st.session_state.get("preorder_orders_local_cache") or {}
    else:
        st.session_state.pop("preorder_orders_local_cache", None)
        st.session_state.pop("preorder_orders_local_key", None)
        loaded_orders = drive_loaded

    drive_name = drive.get("name") or "（無）"
    drive_time = format_gdrive_time(drive.get("modified")) if drive.get("modified") else (
        drive.get("reason") or "❌ 雲端檔案尚未建立/不存在"
    )
    st.caption(f"雲端最新 All Orders：`{drive_name}`　最後修改：`{drive_time}`")

    if loaded_orders.get("preview"):
        st.success(f"目前用本機檔預覽：`{loaded_orders.get('name')}`（未寫入 Drive）")

    if not loaded_orders.get("ok"):
        st.error(f"❌ 無法載入 All Orders：{loaded_orders.get('reason') or '未知錯誤'}")
        st.info("可改本機上傳樣本預覽。請確認 service account 對 Sitegiant_Preorder_Orders 有檢視權，且檔名是 Orders_DD-MM-YYYY-*.xlsx。")
        st.session_state["preorder_orders_synced"] = False
        sku_status = ensure_preorder_sku_status_df(st.session_state.get("preorder_sku_status_df"))
        pending = st.session_state.get("preorder_pending_df")
        sku_status = _render_preorder_sku_and_pending(sku_status, pending, campaign_df, editable=True)
        st.session_state["preorder_sku_status_df"] = sku_status
        return None

    source_label = "本機預覽" if loaded_orders.get("preview") else "雲端"
    modified_label = (
        format_gdrive_time(loaded_orders.get("modified"))
        if loaded_orders.get("modified")
        else "（本機）"
    )
    orders_df = loaded_orders.get("orders")
    order_n = 0 if orders_df is None else len(orders_df)
    preorder_n = 0
    if orders_df is not None and len(orders_df) and "是預購" in orders_df.columns:
        preorder_n = int(orders_df["是預購"].sum())
    st.caption(
        f"看板來源：{source_label} `{loaded_orders.get('name')}`　"
        f"最後修改：`{modified_label}`　"
        f"列數：`{order_n}`　預購列數：`{preorder_n}`"
    )

    pending = extract_preorder_pending_df(orders_df)
    pending = attach_preorder_notify_days(pending, st.session_state.get("preorder_notify_df"))
    sku_status = upsert_preorder_sku_status(
        st.session_state.get("preorder_sku_status_df"),
        pending,
        loaded_orders.get("name") or "",
    )
    sku_key = tuple(sorted(_preorder_text(v) for v in sku_status["SKU"].tolist() if _preorder_text(v)))
    if st.session_state.get("preorder_sku_editor_skus") != sku_key:
        st.session_state.pop("preorder_sku_status_editor", None)
        st.session_state["preorder_sku_editor_skus"] = sku_key
    st.session_state["preorder_pending_df"] = pending
    st.session_state["preorder_sku_status_df"] = sku_status
    st.session_state["preorder_orders_synced"] = True

    sku_status = _render_preorder_sku_and_pending(sku_status, pending, campaign_df, editable=True)
    st.session_state["preorder_sku_status_df"] = sku_status

    board = build_preorder_board(campaign_df, orders_df)
    filled_skus = board.get("campaign_skus") or []
    matched = int(board.get("matched_order_rows") or 0)
    if order_n and preorder_n == 0:
        st.info("本檔沒有商品名稱含「預購」的列，已接單與兩欄為 0。")
    elif preorder_n and not filled_skus:
        st.warning("活動表沒有填 SKU，兩欄對不到預購列。請把上方 SKU 填成預購列的庫存SKU。")
    elif preorder_n and filled_skus and matched == 0:
        shown = "、".join(f"`{s}`" for s in filled_skus[:8])
        st.warning(f"活動 SKU（{shown}）對不到本檔任何預購列，所以兩欄件數是 0。")

    if board["summary"].empty:
        st.info("活動表沒有列，活動看板為空。請先在上方補 SKU。")
        return orders_df

    over = board.get("over_limit_skus") or []
    if over:
        st.warning("⚠️ 已接單已達或超過上限：" + "、".join(f"`{sku}`" for sku in over))

    st.dataframe(board["summary"], use_container_width=True, hide_index=True)
    st.caption("已付款／未付款依付款狀態。可打單另要求訂單狀態＝待處理。催款文只含已標到貨的 SKU。")

    for item in board["campaigns"]:
        sku = item["sku"] or "（未填 SKU）"
        title = (
            f"{sku}　已接單 {item['accepted_qty']}／上限 {item['limit_label']}"
            f"　已付款 {item['paid_qty']}　未付款 {item['unpaid_qty']}"
        )
        if item["over_limit"]:
            title += "　⚠️ 達上限"
        with st.expander(title):
            if item["name"] or item["month"]:
                st.caption(f"月份：{item['month'] or '—'}　品名：{item['name'] or '—'}")
            c_paid, c_unpaid = st.columns(2)
            with c_paid:
                st.markdown("**已付款**")
                st.caption("件數＝商品數量")
                st.metric("已付款件數", item["paid_qty"])
                st.dataframe(item["paid_lines"], use_container_width=True, hide_index=True)
            with c_unpaid:
                st.markdown("**未付款**")
                st.caption("到貨後才催款")
                st.metric("未付款件數", item["unpaid_qty"])
                st.dataframe(item["unpaid_lines"], use_container_width=True, hide_index=True)
    return orders_df


def _render_preorder_sku_and_pending(sku_status, pending, campaign_df, editable=True):
    st.markdown("**SKU 狀態**")
    st.caption("從 All Orders 預購列自動補尚未列過的 SKU，預設預購。標到貨後才進催款文；再同步不會蓋掉到貨。")
    sku_status = ensure_preorder_sku_status_df(sku_status)
    missing = arrived_skus_missing_from_campaign(sku_status, campaign_df)
    if missing:
        st.warning("到貨 SKU 不在活動表（不自動寫入）：" + "、".join(f"`{s}`" for s in missing))
    if editable:
        edited = st.data_editor(
            sku_status,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            column_order=list(PREORDER_SKU_STATUS_COLUMNS),
            column_config={
                "SKU": st.column_config.TextColumn("SKU", disabled=True),
                "品名": st.column_config.TextColumn("品名", disabled=True),
                "件數": st.column_config.NumberColumn("件數", disabled=True),
                "狀態": st.column_config.SelectboxColumn(
                    "狀態",
                    options=[PREORDER_STATUS_PREORDER, PREORDER_STATUS_ARRIVED],
                    required=True,
                ),
                "來源檔": st.column_config.TextColumn("來源檔", disabled=True),
            },
            key="preorder_sku_status_editor",
        )
        sku_status = ensure_preorder_sku_status_df(edited)
    else:
        st.dataframe(sku_status, use_container_width=True, hide_index=True)
    pending_n = 0 if pending is None or getattr(pending, "empty", True) else len(pending)
    with st.expander(f"預購訂單明細（{pending_n} 列）", expanded=False):
        st.caption("最新 All Orders 快照。已過天數只算未付款、未取消、未退款。")
        if pending is None or getattr(pending, "empty", True):
            st.info("目前沒有預購訂單列。")
        else:
            st.dataframe(pending, use_container_width=True, hide_index=True)
    return sku_status


def _render_preorder_arrival_copy(campaign_df, orders_df, sku_status_df):
    st.markdown("### 5. 到貨催款")
    st.caption(
        "只含已標到貨的 SKU。"
        " 催款文＝未付款且非已取消／已退款。"
        " 可打單＝已付款且訂單狀態＝待處理。"
        " 打單、取消、勾已收到付款仍在 SiteGiant。"
    )
    if orders_df is None:
        st.info("請先在「3. 對客戶訂單」載入 All Orders。")
        return
    arrived_campaign = campaign_df_from_arrived_skus(sku_status_df)
    board = build_preorder_board(arrived_campaign, orders_df)
    copies = build_preorder_arrival_copy(board)
    unpaid_text = copies.get("unpaid_reminder") or ""
    paid_text = copies.get("paid_pick") or ""
    st.text_area(
        f"催款可複貼文（{copies.get('unpaid_n', 0)} 列）",
        value=unpaid_text,
        height=220,
        key=f"preorder_copy_unpaid_{hash(unpaid_text)}",
    )
    st.text_area(
        f"可打單名單（待處理＋已付款，{copies.get('paid_n', 0)} 列）",
        value=paid_text,
        height=160,
        key=f"preorder_copy_paid_{hash(paid_text)}",
    )
    if st.button("📌 本次催款已通知", use_container_width=True, key="preorder_stamp_notify"):
        keys = preorder_notify_keys_from_lines(copies.get("unpaid_lines"))
        result = stamp_preorder_first_notify(st.session_state.get("preorder_notify_df"), keys)
        st.session_state["preorder_notify_df"] = result["notify"]
        pending = st.session_state.get("preorder_pending_df")
        st.session_state["preorder_pending_df"] = attach_preorder_notify_days(
            pending, result["notify"]
        )
        if result.get("added"):
            st.session_state["preorder_notify_msg"] = (
                f"已為 {result['added']} 列點上首次通知日 {result.get('today')}（已有日期不改）。"
            )
        else:
            st.session_state["preorder_notify_msg"] = "沒有新的未付款列可點日期（可能已點過，或沒有到貨未付款單）。"
        st.rerun()
    notify_msg = st.session_state.pop("preorder_notify_msg", None)
    if notify_msg:
        st.success(notify_msg)
    st.caption("已過天數給你決定要不要去 SiteGiant 取消訂單；本頁不取消。")


def _render_preorder_vendor_import(campaign_df):
    st.markdown("### 1. 從廠商單帶入")
    st.caption(
        "勾選要跟的款再寫入。只帶入廠商檔名、貨號、品名、條碼；不會填庫存 SKU。"
        " 沒條碼只要勾選也會匯入。已有列的 SKU、自購、上限會保留。"
    )
    uploaded = st.file_uploader(
        "上傳廠商原始訂購單（xlsx）",
        type=["xlsx", "xls"],
        key="preorder_vendor_po_upload",
        help="測用如 2026-License(SJ0910).xlsx。xls 讀失敗請另存 xlsx。",
    )
    if uploaded is None:
        cached = st.session_state.get("preorder_vendor_po")
        if cached:
            st.caption(f"已快取廠商檔：`{cached.get('name')}`（結單回填可用）")
        return campaign_df

    payload = uploaded.getvalue()
    prev = st.session_state.get("preorder_vendor_po") or {}
    if prev.get("name") != uploaded.name or prev.get("bytes") != payload:
        st.session_state.pop("preorder_vendor_fill", None)
        st.session_state.pop("preorder_vendor_preview_editor", None)
    parsed = parse_vendor_po_bytes(payload, uploaded.name)
    st.session_state["preorder_vendor_po"] = {
        "name": uploaded.name,
        "bytes": payload,
        "parsed": parsed,
    }
    if not parsed.get("ok"):
        st.error(f"❌ {parsed.get('reason') or '無法解析廠商檔'}")
        return campaign_df

    preview = parsed.get("preview")
    st.caption(f"檔名：`{uploaded.name}`　可勾選 {len(preview)} 列")
    edited_preview = st.data_editor(
        preview,
        hide_index=True,
        use_container_width=True,
        disabled=["貨號", "品名", "條碼", "sheet", "row"],
        column_config={
            "勾選": st.column_config.CheckboxColumn("勾選", default=True),
        },
        key="preorder_vendor_preview_editor",
    )
    if st.button("⬇️ 把勾選列寫入活動表", use_container_width=True, key="preorder_vendor_import_btn"):
        selected = []
        for rec in edited_preview.to_dict(orient="records"):
            if not rec.get("勾選"):
                continue
            selected.append(rec)
        if not selected:
            st.warning("沒有勾選任何列。")
            return campaign_df
        result = upsert_vendor_rows_into_campaign(campaign_df, selected, uploaded.name)
        st.session_state["preorder_campaign_df"] = result["campaign"]
        _clear_preorder_campaign_editor()
        st.session_state.pop("preorder_lock_preview", None)
        st.session_state.pop("preorder_lock_result", None)
        st.session_state["preorder_vendor_import_msg"] = (
            f"已寫入活動表：新增 {result['added']} 列、更新 {result['updated']} 列。"
        )
        for note in result.get("reports") or []:
            st.session_state["preorder_vendor_import_msg"] += " " + note
        st.rerun()
    return campaign_df


def _render_preorder_close(campaign_df, orders_df):
    campaign_df = st.session_state.get("preorder_campaign_df", campaign_df)
    st.markdown("### 4. 結單下載")
    st.caption(
        "鎖定數量 = 自購 + min(客戶量, 上限)。上限空＝不封頂；自購空＝0。"
        " 未填 SKU 的列不算客戶量、不鎖定。SiteGiant 採購單只給本機下載，不會覆寫雲端空殼。"
        " 同場再確認鎖定，會取代本次尚未存檔的廠商單歷史。"
    )
    dups = duplicate_preorder_skus(campaign_df)
    if dups:
        st.warning("同一 SKU 出現在多列，客戶量會重複加總：" + "、".join(f"`{s}`" for s in dups))

    fill_close = st.checkbox("鎖定時把「實際關閉」填成今天", value=True, key="preorder_fill_close_date")
    close_date = taipei_now().strftime("%Y-%m-%d") if fill_close else None
    lock_result = st.session_state.get("preorder_lock_result")
    preview = st.session_state.get("preorder_lock_preview")

    if st.button("預覽鎖定數量", use_container_width=True, key="preorder_lock_preview_btn"):
        st.session_state["preorder_lock_preview"] = lock_preorder_campaign(
            campaign_df, orders_df, close_date=close_date
        )
        st.session_state.pop("preorder_lock_result", None)
        st.session_state.pop("preorder_vendor_fill", None)
        st.session_state.pop("preorder_restock_fill", None)
        st.rerun()

    if preview and not lock_result:
        st.info("以下為預覽，尚未寫入廠商單歷史。核對後再按確認鎖定。")
        _show_preorder_lock_tables(preview)
        if st.button("確認鎖定", type="primary", use_container_width=True, key="preorder_lock_confirm_btn"):
            result = lock_preorder_campaign(campaign_df, orders_df, close_date=close_date)
            _apply_preorder_lock_to_session(result)
            st.rerun()
        return

    if not lock_result:
        return

    locked_rows = lock_result.get("locked_rows") or []
    st.success(
        f"已鎖定 {len(locked_rows)} 列（{lock_result.get('lock_time')}）。"
        "請按下方「覆寫雲端預購追蹤」把歷史寫回 Drive。"
    )
    _show_preorder_lock_tables(lock_result)

    vendor_cache = st.session_state.get("preorder_vendor_po") or {}
    vendor_bytes = vendor_cache.get("bytes")
    vendor_name = vendor_cache.get("name") or "vendor.xlsx"
    if vendor_bytes:
        st.caption(f"回填使用已上傳的廠商檔：`{vendor_name}`")
    else:
        st.caption("尚未快取廠商檔。請在「1. 從廠商單帶入」再上傳同一份原檔，或下面補傳。")
        extra = st.file_uploader(
            "結單回填用廠商原檔",
            type=["xlsx", "xls"],
            key="preorder_vendor_po_fill_upload",
        )
        if extra is not None:
            vendor_bytes = extra.getvalue()
            vendor_name = extra.name
            st.session_state.pop("preorder_vendor_fill", None)
            st.session_state["preorder_vendor_po"] = {
                "name": vendor_name,
                "bytes": vendor_bytes,
            }

    fill_col, restock_col = st.columns(2)
    with fill_col:
        if vendor_bytes and locked_rows:
            filled = st.session_state.get("preorder_vendor_fill")
            if filled is None:
                filled = fill_vendor_po_qty_bytes(vendor_bytes, locked_rows, vendor_name)
                st.session_state["preorder_vendor_fill"] = filled
            if filled.get("ok"):
                st.download_button(
                    label=f"📥 下載已填訂量的廠商檔（{filled.get('filled', 0)} 列）",
                    data=filled.get("bytes") or b"",
                    file_name=filled.get("filename") or "vendor_qty.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="preorder_vendor_qty_dl",
                )
                unmatched = (filled.get("vendor_unmatched") or []) + (filled.get("campaign_unmatched") or [])
                if unmatched:
                    st.warning("有對不上的列，沒有填假數量。")
                    st.dataframe(pd.DataFrame(unmatched), hide_index=True, use_container_width=True)
            else:
                st.error(f"❌ 回填廠商檔失敗：{filled.get('reason')}")
        elif locked_rows:
            st.info("請上傳廠商原檔才能回填訂量。")

    with restock_col:
        restock = probe_sg_restock_template()
        if not restock.get("ok"):
            st.error(f"❌ 無法讀取 SiteGiant 採購單空殼：{restock.get('reason')}")
        elif locked_rows:
            filled_r = st.session_state.get("preorder_restock_fill")
            if filled_r is None:
                filled_r = fill_sg_restock_bytes(
                    restock.get("bytes"),
                    locked_rows,
                    restock.get("name") or "import_restock.xlsx",
                )
                st.session_state["preorder_restock_fill"] = filled_r
            if filled_r.get("ok"):
                st.download_button(
                    label=f"📥 下載 SiteGiant 採購單（{filled_r.get('rows', 0)} 列，未寫回雲端空殼）",
                    data=filled_r.get("bytes") or b"",
                    file_name=filled_r.get("filename") or "sitegiant採購單.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="preorder_restock_filled_dl",
                )
            else:
                st.error(f"❌ 填 SiteGiant 採購單失敗：{filled_r.get('reason')}")

    _save_preorder_tracker_button(
        st.session_state.get("preorder_campaign_df", campaign_df),
        "preorder_save_lock",
        primary=True,
    )

    hist = st.session_state.get("preorder_vendor_history_df")
    if hist is not None and len(hist):
        with st.expander(f"廠商單歷史（{len(hist)} 列）"):
            show = hist.copy()
            cols = [c for c in PREORDER_VENDOR_HISTORY_COLUMNS if c in show.columns]
            extras = [c for c in show.columns if c not in cols]
            st.dataframe(show[cols + extras], hide_index=True, use_container_width=True)


def _xlsx_filename(name):
    """openpyxl 產出的是 xlsx 位元組，上傳 MIME 必須跟副檔名一致，避免把 OOXML 標成 .xls。"""
    raw = str(name).strip()
    lower = raw.lower()
    if lower.endswith(".xlsx"):
        return raw
    if lower.endswith(".xls"):
        return raw[:-4] + ".xlsx"
    if "." in raw:
        return raw.rsplit(".", 1)[0] + ".xlsx"
    return f"{raw}.xlsx"


def _read_inward_excel(file_bytes):
    payload = file_bytes.getvalue() if isinstance(file_bytes, io.BytesIO) else file_bytes
    engine_kw = {"engine": "calamine"} if HAS_CALAMINE else {}
    return pd.read_excel(io.BytesIO(payload), **engine_kw)


def _hist_fix_error_text(exc):
    msg = str(exc)
    lowered = msg.lower()
    if any(k in lowered for k in ["憑證", "credential", "private_key", "textkey", "service_account"]):
        return "憑證缺失／無效"
    if "欄位" in msg or "找不到" in msg:
        return msg
    return "讀寫失敗"


def _show_hist_fix_results(results, key_prefix):
    if not results:
        return
    for i, item in enumerate(results):
        name = item["name"]
        if item.get("error"):
            st.error(f"❌ `{name}`：{item['error']}")
            continue
        updated = item.get("updated", 0)
        not_found = item.get("not_found") or []
        empty_sku = item.get("empty_sku") or []
        status = item.get("status")
        if status == "overwritten":
            st.success(f"✅ `{name}` 已覆寫雲端（更新 {updated} 筆庫存貨品名稱）")
        elif status == "download_only":
            st.warning(
                f"⚠️ `{name}` 雲端沒有同名檔，未寫入。請先在歷史入庫資料夾手動建立該檔。"
                f"已更新 {updated} 筆名稱，可先下載。"
            )
        else:
            st.info(f"📄 `{name}`：更新 {updated} 筆庫存貨品名稱")
        if not_found:
            with st.expander(f"找不到統整表對應（{len(not_found)} 筆）— {name}"):
                st.write(not_found)
        if empty_sku:
            with st.expander(f"統整表 sitegiant庫存SKU 空白（{len(empty_sku)} 筆）— {name}"):
                st.write(empty_sku)
        if item.get("bytes"):
            dl_name = _xlsx_filename(name)
            st.download_button(
                label=f"📥 下載更正後檔案：{dl_name}",
                data=item["bytes"],
                file_name=dl_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"{key_prefix}_dl_{i}_{name}",
            )


def _fix_inward_df_from_summary(df, name_map, empty_sku_barcodes):
    updated_df, summary = apply_sitegiant_name_from_summary(df, name_map, empty_sku_barcodes)
    return updated_df, summary, _inward_excel_bytes(updated_df)


def render(sub_page, ID_PRICE_SUMMARY, ID_HISTORY_INWARD_FOLDER, ID_SHOPEE_MASTER, ID_PRICE_SUMMARY_FALLBACK=None, ID_SITEGIANT_UPC_FOLDER=None):
    """
    渲染 Sitegiant 電商整合管理的所有子頁面
    將需要的 Google Drive ID 當作參數傳進來
    """
    
    st.title(f"{sub_page}")
    st.info(f"目前導覽路徑： 🌐 Sitegiant 電商整合管理 ➔ {sub_page}")
    st.write("---")

    # -------------------------------------------------------------------------
    # 子功能 1：🔀 sitegiant 採購入庫單格式轉換
    # -------------------------------------------------------------------------
    if sub_page == "🔀 Sitegiant 採購入庫單格式轉換":
        st.subheader("🛍️ Sitegiant 採購入庫單內容填寫")
        
        # ── 1. 基本設定與資料輸入 ──
        c_meta1, c_meta2 = st.columns(2)
        with c_meta1: 
            order_no = st.text_input("📝 請輸入訂單/銷貨單號：", value=datetime.date.today().strftime("%Y%m%d01"))
            vendor_options = ["麗嬰", "Buyee", "日亞", "其他"]
            selected_vendor = st.selectbox("🏬 請選擇採購廠商：", vendor_options, key="sg_vendor_selectbox")
            
            if selected_vendor == "其他":
                custom_vendor = st.text_input("✍️ 請輸入自訂廠商名稱：", key="sg_custom_vendor_name")
                vendor_name = custom_vendor if custom_vendor.strip() else "其他廠商"
            else:
                vendor_name = selected_vendor

        with c_meta2:
            recv_date = st.date_input("📅 選擇銷貨日期：", value=datetime.date.today())
            recv_date = recv_date.strftime("%y%m%d")
            
        st.write("---")

        # ── 2. 剪貼簿快速文字貼上區 ──
        st.markdown("### 1️⃣ 第一步：貼上原始資料")
        st.info("💡 請直接從 Excel 複製『國際條碼』與『數量』這兩欄資料，並貼入下方文字框中。")
        
        pasted_text = st.text_area(
            "📋 剪貼簿貼上區：", 
            height=150, 
            placeholder="請在此貼上...\n範例格式：\n4711234567890\t2\n4711234567891\t5"
        )
        
        if st.button("📥 解析並產生預覽表格", type="secondary", use_container_width=True):
            if pasted_text.strip():
                try:
                    df_parsed = pd.read_csv(io.StringIO(pasted_text.strip()), sep=r'\s+|\t', engine='python', header=None, dtype=str)
                    col_count = df_parsed.shape[1]
                    
                    if col_count >= 2:
                        df_parsed = df_parsed.iloc[:, :2]
                        df_parsed.columns = ["國際條碼", "數量"]
                    elif col_count == 1:
                        df_parsed = df_parsed.iloc[:, :1]
                        df_parsed.columns = ["國際條碼"]
                        df_parsed["數量"] = "1"
                    else:
                        raise ValueError("文字框內容為空或格式無法辨識。")
                    
                    st.session_state['inward_input_df'] = df_parsed
                    st.success(f"✅ 成功解析 {len(df_parsed)} 筆資料！已同步至下方預覽表格。")
                    import time
                    time.sleep(1.0)
                    st.rerun()
                except Exception as parse_err:
                    st.error(f"❌ 解析失敗，請確認複製內容格式是否正確：{parse_err}")
            else:
                st.warning("⚠️ 文字框為空，請先貼上資料。")

        st.write("---")

        # ── 3. 動態預覽與編輯區 ──
        st.markdown("### 2️⃣ 第二步：檢查與微調明細")
        if 'inward_input_df' not in st.session_state:
            st.session_state['inward_input_df'] = pd.DataFrame(columns=["國際條碼", "數量"])
            
        input_df = st.data_editor(
            st.session_state['inward_input_df'], 
            num_rows="dynamic", 
            use_container_width=True, 
            key="inward_grid"
        )
        st.session_state['inward_input_df'] = input_df 

        st.write("---") 
        
        # ── 4. 核心執行按鈕 ──
        st.markdown("### 3️⃣ 第三步：執行格式轉換")
        if st.button("✨ 執行貨品名稱和成本稅款導入並紀錄待處理商品", type="primary", use_container_width=True):
            if not order_no.strip() or input_df.empty: 
                st.error("❌ 轉換失敗！請填入銷貨單號，並確認上方表格有有效明細。")
            else:
                with st.spinner("正在由雲端獲取最新商品統整表並進行精準關聯..."):
                    try:
                        if ID_PRICE_SUMMARY:
                            engine_kw = {"engine": "calamine"} if HAS_CALAMINE else {}
                            df_ref = pd.read_excel(download_gdrive_file_to_bytes(ID_PRICE_SUMMARY), **engine_kw)
                            df_ref['c_clean'] = df_ref['c'].astype(str).str.strip().str.split('.').str[0]
                                
                            result_rows = []
                            missing_items = [] 
                            
                            for row in input_df.itertuples(index=False):
                                barcode_input = str(row.國際條碼).strip().split('.')[0] if pd.notna(row.國際條碼) else ""
                                if barcode_input in ["", "0", "nan", "None"]: continue
                                qty = int(row.數量) if pd.notna(row.數量) else 0
                            
                                sku_final = "⚠️ 提示：須新增iSKU"
                                prod_name = "⚠️ 未知商品名稱" 
                                category = ""
                                keywords = ""
                                cost_val = None 
                                tax_val = None
                                or_val = None
                                sal_val = None
                            
                                if not df_ref.empty and 'c_clean' in df_ref.columns:
                                    match = df_ref[df_ref['c_clean'] == barcode_input]
                                    if not match.empty:
                                        match_row = match.iloc[0]
                                        
                                        sg_name = match_row.get('sitegiant庫存SKU', None)
                                        if pd.notna(sg_name) and str(sg_name).strip() not in ["", "nan", "None"]:
                                            prod_name = str(sg_name).strip()
                                        
                                        sku = match_row.get('自定義編碼', '')
                                        if pd.notna(sku) and str(sku).strip() != "":
                                            sku_final = str(sku).strip()
                                        
                                        category = match_row.get('分類定義', '')
                                        keywords = match_row.get('產品關鍵字', '')
                                    
                                        if vendor_name == "麗嬰":
                                            def safe_float(v):
                                                try:
                                                    return float(str(v).replace(',', '').strip())
                                                except (ValueError, TypeError):
                                                    return None

                                            c_val = safe_float(match_row.get('麗嬰未稅價', None))
                                            t_val = safe_float(match_row.get('麗嬰稅款', None))
                                            s_val = safe_float(match_row.get('麗嬰批發含稅價', None))
                                            
                                            cost_val = round(c_val, 2) if c_val is not None else None
                                            tax_val = round(t_val, 2) if t_val is not None else None
                                            or_val = safe_float(match_row.get('麗嬰零售價', None))
                                            sal_val = round(s_val, 2) if s_val is not None else None

                                issue_type = ""
                                if sku_final == "⚠️ 提示：須新增iSKU":
                                    issue_type = "須建立自定義編碼"
                                elif prod_name == "⚠️ 未知商品名稱":
                                    issue_type = "須建立賣場商品"
                                    
                                if issue_type:
                                    current_order_name = order_no if 'order_no' in locals() and str(order_no).strip() else "手動輸入未命名單號"
                                    missing_items.append({
                                        "採購單檔名": current_order_name,
                                        "國際條碼": str(barcode_input).strip(),
                                        "狀況": issue_type,  
                                        "狀態": "待處理",    
                                        "建立時間": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                    })

                                result_rows.append({
                                    "銷貨日期": str(recv_date), "國際條碼": barcode_input,
                                    "庫存SKU": sku_final, "庫存貨品名稱": prod_name, 
                                    "麗嬰零售價": or_val if vendor_name == "麗嬰" else None, 
                                    "麗嬰批發含稅價": sal_val if vendor_name == "麗嬰" else None,
                                    "成本": cost_val, "稅款": tax_val, "數量": qty,
                                    "分類定義": category, "產品關鍵字": keywords
                                })
                                
                            # =========================================================
                            # 🔄 同步更新「尚未建立商品清單」(包含自動清理與新增)
                            # =========================================================
                            try:
                                TARGET_SHEET_ID = "1Ixp9V_u2yU8hiWhxQCHNDB4kPKlxDGD2"
                                
                                # 1. 取得雲端最新清單
                                try:
                                    raw_bytes = download_gdrive_file_to_bytes(TARGET_SHEET_ID)
                                    engine_kw = {"engine": "calamine"} if HAS_CALAMINE else {}
                                    df_missing = pd.read_excel(raw_bytes, sheet_name=0, dtype=str, **engine_kw)
                                except Exception:
                                    df_missing = pd.DataFrame(columns=["採購單檔名", "國際條碼", "狀況", "狀態", "建立時間"])
                                    
                                df_missing.columns = df_missing.columns.astype(str).str.strip()
                                changed = False  
                                
                                # 2. 🧹 執行同步移除
                                if "國際條碼" in df_missing.columns and "狀態" in df_missing.columns:
                                    input_barcodes = set(clean_barcode(b) for b in input_df['國際條碼'].tolist() if pd.notna(b))
                                    current_missing_barcodes = set(clean_barcode(item['國際條碼']) for item in missing_items)
                                    df_missing['_clean_barcode'] = df_missing['國際條碼'].apply(clean_barcode)
                                    
                                    cond_auto_resolved = df_missing['_clean_barcode'].apply(
                                        lambda x: (x in input_barcodes) and (x not in current_missing_barcodes)
                                    )
                                    
                                    status_col = df_missing['狀態'].astype(str).str.strip()
                                    cond_manual_resolved = (status_col != '待處理') & (~status_col.isin(['nan', 'None', '', 'NaN']))
                                    
                                    mask_to_remove = cond_auto_resolved | cond_manual_resolved
                                    removed_count = mask_to_remove.sum()
                                    
                                    if removed_count > 0:
                                        df_missing = df_missing[~mask_to_remove] 
                                        st.toast(f"🧹 已同步從清單中自動清理 {removed_count} 筆已解決的商品！", icon="✅")
                                        changed = True
                                        
                                    df_missing = df_missing.drop(columns=['_clean_barcode'])
                                
                                # 3. 🚨 執行新增
                                if missing_items:
                                    if "採購單檔名" in df_missing.columns and "國際條碼" in df_missing.columns:
                                        existing_keys = set(df_missing["採購單檔名"].astype(str).str.strip() + "_" + df_missing["國際條碼"].astype(str).str.strip())
                                    else:
                                        existing_keys = set()
                                    
                                    new_missing_items = []
                                    for item in missing_items:
                                        item_key = f"{item['採購單檔名']}_{item['國際條碼']}"
                                        if item_key not in existing_keys:
                                            new_missing_items.append(item)
                                            existing_keys.add(item_key)
                                    
                                    if new_missing_items:
                                        df_new_rows = pd.DataFrame(new_missing_items)
                                        df_missing = pd.concat([df_missing, df_new_rows], ignore_index=True)
                                        st.toast(f"🚨 已自動將 {len(new_missing_items)} 筆異常紀錄新增至待處理清單！", icon="⚠️")
                                        changed = True
                                
                                # 4. 💾 上傳回 Google Drive
                                if changed:
                                    output_stream = io.BytesIO()
                                    with pd.ExcelWriter(output_stream, engine='openpyxl') as writer:
                                        df_missing.to_excel(writer, index=False, sheet_name="尚未建立商品清單")
                                    output_stream.seek(0)
                                    
                                    upload_or_update_gdrive_file(
                                        folder_id=None,
                                        file_name="尚未建立商品清單.xlsx", 
                                        file_bytes=output_stream.getvalue(),
                                        existing_file_id=TARGET_SHEET_ID
                                    )
                            except Exception as log_err:
                                st.error(f"⚠️ 同步更新「尚未建立商品清單」失敗: {str(log_err)}")   
                                    
                            if result_rows:
                                st.session_state['inward_result_df'] = pd.DataFrame(result_rows)
                                st.session_state['current_vendor_name'] = vendor_name
                                st.session_state['current_order_no'] = order_no
                                st.session_state['has_pending_items'] = len(missing_items) > 0 
                                st.success(f"🚀 格式勾稽完成！廠商已設定為：【{vendor_name}】")
                                st.rerun()
                                
                        else:
                            st.error("❌ 雲端找不到『商品蝦皮麗嬰價格統整表』。")
                            
                    except Exception as e: 
                        st.error(f"❌ 錯誤: {str(e)}")

        # ── 5. 結果預覽與下載區 ──
        if 'inward_result_df' in st.session_state:
            st.write("---")
            res_df = st.session_state['inward_result_df']
            current_vendor = st.session_state.get('current_vendor_name', '未命名廠商')
            current_order = st.session_state.get('current_order_no', '0000')
            
            target_columns = ["國際條碼","庫存SKU", "庫存貨品名稱", "麗嬰零售價", "麗嬰批發含稅價", "成本", "稅款", "數量"]
            available_cols = [col for col in target_columns if col in res_df.columns]
            df_download = res_df[available_cols].copy()
            
            st.markdown(f"### 📋 【{current_vendor}】入庫明細結果預覽")
            
            edited_inward_df = st.data_editor(
                df_download,
                use_container_width=True,
                disabled=["國際條碼","庫存SKU", "庫存貨品名稱", "數量", "麗嬰零售價"],
                column_config={
                    "麗嬰批發含稅價": st.column_config.NumberColumn("麗嬰批發含稅價", min_value=0.0, format="%.2f"),
                    "成本": st.column_config.NumberColumn("成本", min_value=0.0, format="%.2f"),
                    "稅款": st.column_config.NumberColumn("稅款", min_value=0.0, format="%.2f"),
                },
                key="inward_items_editor_final"
            )
            
            h_cost = 0.0
            h_tax = 0.0
            for h_row in edited_inward_df.itertuples(index=False):
                h_qty = int(h_row.數量) if (hasattr(h_row, '數量') and pd.notna(h_row.數量)) else 0
                if hasattr(h_row, '成本') and pd.notna(h_row.成本):
                    try: h_cost += float(h_row.成本) * h_qty
                    except: pass
                if hasattr(h_row, '稅款') and pd.notna(h_row.稅款):
                    try: h_tax += float(h_row.稅款) * h_qty
                    except: pass
            
            st.markdown("#### 📊 本張單據入庫成本稅款即時統計看板")
            c_tot1, c_tot2 = st.columns(2)
            with c_tot1: 
                st.metric(label="💰 當前單據成本未稅總金額 (成本 * 數量)", value=f"$ {h_cost:,.2f} 元")
            with c_tot2: 
                st.metric(label="🧾 當前單據營業稅總金額 (稅款 * 數量)", value=f"$ {h_tax:,.2f} 元")
            
            st.write("---")
            
            col_dl, col_reset = st.columns(2)
            with col_dl:
                towrite_inward = io.BytesIO()
                with pd.ExcelWriter(towrite_inward, engine='openpyxl') as writer:
                    edited_inward_df.to_excel(writer, index=False, sheet_name="SiteGiant入庫單")
                    
                has_pending = st.session_state.get('has_pending_items', False)
                pending_suffix = "_待處理" if has_pending else ""
                final_filename = f"sitegiant採購入庫單_{recv_date}_{current_vendor}_{current_order}{pending_suffix}.xlsx"

                st.download_button(
                    label=f"📥 儲存並下載 Sitegiant 格式入庫單",
                    data=towrite_inward.getvalue(),
                    file_name=final_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
                
            with col_reset:
                if st.button("🔄 完成下載，清除畫面處理下一筆", type="secondary", use_container_width=True):
                    if 'inward_result_df' in st.session_state:
                        del st.session_state['inward_result_df']
                    st.session_state['inward_input_df'] = pd.DataFrame(columns=["國際條碼", "數量"])
                    st.rerun()

    # -------------------------------------------------------------------------
    # 子功能 2：📜 Sitegiant 歷史入庫單紀錄
    # -------------------------------------------------------------------------
    elif sub_page == "📜 Sitegiant 歷史入庫單紀錄":
        summary_file_id = ID_PRICE_SUMMARY or ID_PRICE_SUMMARY_FALLBACK
        hist_files = list_history_inward_files(ID_HISTORY_INWARD_FOLDER)
        hist_options = history_inward_option_map(hist_files)

        st.subheader("📥 批次匯入歷史入庫單")
        st.caption(
            "以「國際條碼」對統整表欄位 `c`，將 `sitegiant庫存SKU` 寫入「庫存貨品名稱」。"
            "入庫單在 `YYYY/YYMM` 子資料夾；雲端已有同名檔才覆寫（跨月同名時對檔名日期所在月份）。"
            "沒有同名檔只提供本機下載，不新建。"
        )
        uploaded_hist_files = st.file_uploader(
            "📥 選擇歷史入庫單 Excel（可多選）",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            key="hist_inward_import_files",
        )
        if st.button("✨ 依統整表更新庫存貨品名稱並匯入", type="primary", use_container_width=True, key="hist_inward_import_btn"):
            if not uploaded_hist_files:
                st.error("❌ 請先選擇要匯入的入庫單檔案。")
            else:
                with st.spinner("⏳ 正在對照統整表並處理匯入檔案..."):
                    try:
                        name_map, empty_sku_barcodes = load_barcode_to_sitegiant_name_map(summary_file_id)
                        results = []
                        for uploaded in uploaded_hist_files:
                            file_name = uploaded.name
                            try:
                                df_in = _read_inward_excel(uploaded.getvalue())
                                _, summary, xlsx_bytes = _fix_inward_df_from_summary(
                                    df_in, name_map, empty_sku_barcodes
                                )
                                existing = match_history_inward_file(hist_files, file_name)
                                if existing:
                                    upload_or_update_gdrive_file(
                                        ID_HISTORY_INWARD_FOLDER,
                                        _xlsx_filename(file_name),
                                        xlsx_bytes,
                                        existing_file_id=existing["id"],
                                    )
                                    status = "overwritten"
                                else:
                                    status = "download_only"
                                results.append({
                                    "name": history_inward_file_label(existing) if existing else file_name,
                                    "status": status,
                                    "updated": summary["updated"],
                                    "not_found": summary["not_found"],
                                    "empty_sku": summary["empty_sku"],
                                    "bytes": xlsx_bytes,
                                })
                            except Exception as file_err:
                                results.append({
                                    "name": file_name,
                                    "error": _hist_fix_error_text(file_err),
                                })
                        st.session_state["hist_import_results"] = results
                        if any(item.get("status") == "overwritten" for item in results):
                            invalidate_history_inward_caches()
                            st.session_state.pop("inward_index_result", None)
                    except Exception as e:
                        st.error(f"❌ 無法讀取統整表：{_hist_fix_error_text(e)}")
                        st.session_state["hist_import_results"] = []

        if st.session_state.get("hist_import_results"):
            _show_hist_fix_results(st.session_state["hist_import_results"], "hist_import")

        st.write("---")
        st.subheader("🔄 批次修正雲端既有單據")
        if not hist_files:
            st.warning("💡 目前雲端無歷史單據可修正。")
        else:
            hist_labels = list(hist_options.keys())
            selected_fix_files = st.multiselect(
                "選擇要修正的歷史入庫單：",
                hist_labels,
                default=hist_labels,
                key="hist_inward_fix_select",
            )
            if st.button("✨ 依統整表更新所選單據的庫存貨品名稱", type="primary", use_container_width=True, key="hist_inward_fix_btn"):
                if not selected_fix_files:
                    st.error("❌ 請至少選擇一張歷史入庫單。")
                else:
                    with st.spinner("⏳ 正在對照統整表並覆寫雲端單據..."):
                        try:
                            name_map, empty_sku_barcodes = load_barcode_to_sitegiant_name_map(summary_file_id)
                            results = []
                            for label in selected_fix_files:
                                existing = hist_options.get(label)
                                if not existing:
                                    results.append({"name": label, "error": "雲端找不到該檔案"})
                                    continue
                                try:
                                    file_bytes = download_gdrive_file_to_bytes(existing["id"])
                                    df_in = _read_inward_excel(file_bytes)
                                    _, summary, xlsx_bytes = _fix_inward_df_from_summary(
                                        df_in, name_map, empty_sku_barcodes
                                    )
                                    upload_or_update_gdrive_file(
                                        ID_HISTORY_INWARD_FOLDER,
                                        _xlsx_filename(existing["name"]),
                                        xlsx_bytes,
                                        existing_file_id=existing["id"],
                                    )
                                    results.append({
                                        "name": label,
                                        "status": "overwritten",
                                        "updated": summary["updated"],
                                        "not_found": summary["not_found"],
                                        "empty_sku": summary["empty_sku"],
                                        "bytes": xlsx_bytes,
                                    })
                                except Exception as file_err:
                                    results.append({
                                        "name": label,
                                        "error": _hist_fix_error_text(file_err),
                                    })
                            st.session_state["hist_fix_results"] = results
                            if any(item.get("status") == "overwritten" for item in results):
                                invalidate_history_inward_caches()
                                st.session_state.pop("inward_index_result", None)
                        except Exception as e:
                            st.error(f"❌ 無法讀取統整表：{_hist_fix_error_text(e)}")
                            st.session_state["hist_fix_results"] = []

            if st.session_state.get("hist_fix_results"):
                _show_hist_fix_results(st.session_state["hist_fix_results"], "hist_fix")

        st.write("---")
        st.subheader("📊 歷史入庫單與成本稅款加總指標檢視")
        if not hist_files:
            st.warning("💡 目前雲端無歷史單據。")
        else:
            selected_hist_label = st.selectbox("🎯 選擇欲調閱的入庫對帳單：", list(hist_options.keys()))
            if selected_hist_label:
                try:
                    selected_meta = hist_options[selected_hist_label]
                    target_id = selected_meta["id"]
                    file_bytes = download_gdrive_file_to_bytes(target_id)
                    df_hist_view = pd.read_excel(file_bytes, engine="calamine" if HAS_CALAMINE else None)
                    st.markdown(
                        f"📄 **當前雲端檔案**：`{selected_hist_label}` ｜ "
                        f"📊 **單據品項數**：`{len(df_hist_view)} 筆`"
                    )
                    st.dataframe(df_hist_view, use_container_width=True)
                    
                    h_cost = 0.0
                    h_tax = 0.0
                    for h_row in df_hist_view.itertuples(index=False):
                        h_qty = int(h_row.數量) if (hasattr(h_row, '數量') and pd.notna(h_row.數量)) else 0
                        if hasattr(h_row, '成本') and pd.notna(h_row.成本):
                            try: h_cost += float(h_row.成本) * h_qty
                            except ValueError: pass
                        if hasattr(h_row, '稅款') and pd.notna(h_row.稅款):
                            try: h_tax += float(h_row.稅款) * h_qty
                            except ValueError: pass
                        
                    st.markdown("#### 📊 本張單據入庫成本稅款")
                    c_tot1, c_tot2 = st.columns(2)
                    with c_tot1: 
                        st.metric(label="💰 成本未稅總金額 (成本 * 數量)", value=f"$ {h_cost:,.2f} 元")
                    with c_tot2: 
                        st.metric(label="🧾 營業稅總金額 (稅款 * 數量)", value=f"$ {h_tax:,.2f} 元")    
                    
                    st.download_button(
                        label="🔄 下載此歷史採購入庫單",
                        data=file_bytes.getvalue(),
                        file_name=selected_meta["name"],
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                except Exception as e: st.error(f"❌ 讀取失敗: {str(e)}")

    # -------------------------------------------------------------------------
    # 子功能 2b：🔍 查詢入庫紀錄
    # -------------------------------------------------------------------------
    elif sub_page == "🔍 查詢入庫紀錄":
        st.subheader("🔍 查詢商品來自哪筆歷史入庫單")
        st.caption(
            "掃描歷史入庫根目錄與 `YYYY/YYMM` 子資料夾。"
            f"查詢先讀根目錄 `{HISTORY_INWARD_INDEX_NAME}`；來源清單有增刪改才重掃並覆寫索引。"
        )
        if st.button("🔄 重新載入索引", key="inward_query_reload"):
            st.session_state.pop("inward_index_result", None)
            with st.spinner("⏳ 正在重新載入入庫紀錄索引..."):
                st.session_state["inward_index_result"] = refresh_history_inward_index(ID_HISTORY_INWARD_FOLDER)
            st.rerun()
        if "inward_index_result" not in st.session_state:
            with st.spinner("⏳ 正在載入入庫紀錄索引..."):
                st.session_state["inward_index_result"] = load_history_inward_index(ID_HISTORY_INWARD_FOLDER)
        index_result = st.session_state["inward_index_result"]
        df_index = index_result.get("df")
        if df_index is None:
            df_index = pd.DataFrame()

        if not index_result.get("index_exists"):
            st.warning(
                f"⚠️ 找不到雲端 `{HISTORY_INWARD_INDEX_NAME}`。"
                "本次仍可查詢，結果只留在畫面。"
            )
        elif index_result.get("rebuilt") and index_result.get("written"):
            st.info("ℹ️ 來源入庫單有變動，已重建並覆寫雲端索引。")
        elif index_result.get("rebuilt") and not index_result.get("written"):
            st.warning("⚠️ 已於本次重建索引，但未能寫回雲端。")

        skipped = index_result.get("skipped") or []
        if skipped:
            with st.expander(f"略過無法讀取的檔案（{len(skipped)}）"):
                for item in skipped:
                    st.write(f"- `{item.get('name')}`：{item.get('reason')}")

        c_stat1, c_stat2 = st.columns(2)
        with c_stat1:
            st.metric("入庫單檔數", f"{index_result.get('file_count', 0)}")
        with c_stat2:
            st.metric("索引明細列數", f"{len(df_index)}")

        st.write("---")
        st.markdown("#### 🔎 多筆批次查詢")
        col_m, col_i = st.columns([1, 3])
        with col_m:
            target_col = st.radio("選擇查詢依據欄位：", options=["國際條碼", "庫存SKU"], index=0)
        with col_i:
            batch_input = st.text_area(
                f"請輸入多筆【{target_col}】（每筆請以換行、逗號或空格隔開）：",
                height=100,
            )

        df_display = df_index.copy()
        missing_terms = []
        if batch_input.strip():
            search_terms = [t.strip() for t in re.split(r"[\n,\s]+", batch_input) if t.strip()]
            if target_col == "國際條碼":
                cleaned_terms = [clean_barcode(t) for t in search_terms]
                cleaned_terms = [t for t in cleaned_terms if t]
                if "國際條碼" in df_display.columns:
                    df_display["國際條碼"] = df_display["國際條碼"].map(clean_barcode)
                    df_display = df_display[df_display["國際條碼"].isin(cleaned_terms)]
                    found = set(df_display["國際條碼"].tolist())
                    missing_terms = [t for t in cleaned_terms if t not in found]
            else:
                cleaned_terms = [str(t).strip() for t in search_terms]
                if cleaned_terms and "庫存SKU" in df_display.columns:
                    sku_series = df_display["庫存SKU"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
                    df_display = df_display[sku_series.isin(cleaned_terms)]
                    found = set(sku_series[sku_series.isin(cleaned_terms)].tolist())
                    missing_terms = [t for t in cleaned_terms if t not in found]
            st.info(f"🎯 批次篩選結果：找到 **{len(df_display)}** 筆符合資料。")
            if missing_terms:
                st.warning("查無資料：" + "、".join(missing_terms))

        if not df_display.empty and "銷貨日期" in df_display.columns:
            df_display = df_display.sort_values(
                by=["銷貨日期", "檔名"] if "檔名" in df_display.columns else ["銷貨日期"],
                ascending=[False] * (2 if "檔名" in df_display.columns else 1),
                kind="mergesort",
            ).reset_index(drop=True)

        qty_sum = 0
        if "數量" in df_display.columns:
            qty_sum = pd.to_numeric(df_display["數量"], errors="coerce").fillna(0).sum()
        m1, m2 = st.columns(2)
        with m1:
            st.metric("符合筆數", f"{len(df_display)}")
        with m2:
            st.metric("進貨數量合計", f"{int(qty_sum)}")

        st.dataframe(df_display, use_container_width=True)

        towrite_query = io.BytesIO()
        with pd.ExcelWriter(towrite_query, engine="openpyxl") as writer:
            df_display.to_excel(writer, index=False, sheet_name="入庫查詢結果")
        st.download_button(
            label="📥 下載本次查詢結果 (.xlsx)",
            data=towrite_query.getvalue(),
            file_name=f"入庫紀錄查詢_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # -------------------------------------------------------------------------
    # 子功能 3：📦 SiteGiant 批量新增UPC
    # -------------------------------------------------------------------------
    elif sub_page == "📦 Sitegiant 批量新增UPC":
        st.subheader("📦 Sitegiant 批量新增UPC")
        
        with st.expander("📌 點擊查看操作流程與前/後置條件", expanded=True):
            st.markdown("""
            ### 📝 前置條件
            1. **確認 iSKU 對應 UPC**。
            2. **將 Sitegiant 批量編輯 UPC 檔放到雲端 `Sitegiant_UPC` 資料夾**（xlsx 或 zip 皆可）：
               - 檔名格式：`batch_edit_item_upc_assignment_all_DD-MM-YYYY-*.xlsx`
            3. **確認蝦皮賣場列表已校正**。
            ### 🚀 後續操作
            - 將下方處理完畢的 Excel 上傳回 Sitegiant 覆蓋即可。定時同步也會把結果寫成雲端 `batch_edit_upc_added_only.xlsx`。
            """)

        df_shopee_hist, df_shopee_list = load_shopee_data(ID_SHOPEE_MASTER)

        if df_shopee_list.empty:
            st.error("❌ 無法從雲端讀取蝦皮商品列表！請先前往「蝦皮商品清單轉換」執行校正並回寫雲端。")
        else:
            st.info("✅ 系統已自動從雲端載入最新的蝦皮商品列表，準備好進行 UPC 交叉比對！")

            latest_upc = None
            if ID_SITEGIANT_UPC_FOLDER:
                latest_upc = pick_latest_gdrive_file(
                    ID_SITEGIANT_UPC_FOLDER,
                    "batch_edit_item_upc_assignment_all",
                    "dmy",
                    include_zip=True,
                )
            if latest_upc:
                st.success(
                    f"將處理最新來源檔：`{latest_upc['name']}` ｜ 📅 `{format_gdrive_time(latest_upc.get('modifiedTime'))}`"
                )
                st.caption("寫回 `batch_edit_upc_added_only.xlsx`：請先在資料夾手動建立此檔，service account 只能覆寫、不能新建。")
            else:
                st.error("❌ 資料夾內找不到 `batch_edit_item_upc_assignment_all_DD-MM-YYYY-*.xlsx/.zip`。")

            def _run_upc_fill(file_bytes, source_name, write_drive=False):
                df_sg = read_tabular_file(file_bytes, source_name)
                filled, stats = fill_sitegiant_upc(df_sg, df_shopee_list)
                st.session_state["sg_upc_updated_df"] = filled
                if stats["updated"] > 0:
                    st.success(f"🎉 處理完成！共成功自動填補 **{stats['updated']}** 筆缺失的 UPC 資料。")
                    if write_drive and ID_SITEGIANT_UPC_FOLDER:
                        existing = resolve_named_file(ID_SITEGIANT_UPC_FOLDER, "batch_edit_upc_added_only")
                        towrite_sg = io.BytesIO()
                        with pd.ExcelWriter(towrite_sg, engine="openpyxl") as writer:
                            filled.to_excel(writer, index=False)
                        upload_or_update_gdrive_file(
                            ID_SITEGIANT_UPC_FOLDER,
                            UPC_FILLED_FILENAME,
                            towrite_sg.getvalue(),
                            existing_file_id=existing["id"] if existing else None,
                            allow_create=False,
                        )
                        st.info(f"☁️ 已寫回雲端 `{UPC_FILLED_FILENAME}`，請再上傳至 Sitegiant。")
                else:
                    st.warning("⚠️ 處理完成，但未找到任何可填補的缺失 UPC 資料。")

            if st.button("⚡ 從雲端最新 UPC 檔填補", type="primary", use_container_width=True):
                if not latest_upc:
                    st.error("❌ 沒有可處理的來源檔。")
                else:
                    with st.spinner("⏳ 正在下載雲端檔並填入缺失的 UPC..."):
                        try:
                            payload, inner_name = download_source_spreadsheet(
                                latest_upc, "batch_edit_item_upc_assignment_all"
                            )
                            _run_upc_fill(payload, inner_name, write_drive=True)
                        except Exception as e:
                            st.error(f"❌ 處理檔案時發生錯誤：{str(e)}")

            with st.expander("備用：手動上傳 UPC 檔"):
                uploaded_sg_file = st.file_uploader(
                    "📥 請上傳 Sitegiant UPC 批量編輯下載檔 (.xlsx/.xls/.csv/.zip)",
                    type=["xlsx", "xls", "csv", "zip"],
                )
                if uploaded_sg_file and st.button("⚡ 執行上傳檔自動比對並填補 UPC", type="secondary", use_container_width=True):
                    with st.spinner("⏳ 正在自動比對蝦皮資料庫並填入缺失的 UPC..."):
                        try:
                            raw = uploaded_sg_file.read()
                            name = uploaded_sg_file.name
                            if name.lower().endswith(".zip"):
                                raw, name = extract_xlsx_from_zip(raw, "batch_edit_item_upc_assignment_all")
                            _run_upc_fill(raw, name, write_drive=False)
                        except Exception as e:
                            st.error(f"❌ 處理檔案時發生錯誤：{str(e)}")

            if "sg_upc_updated_df" in st.session_state and not st.session_state["sg_upc_updated_df"].empty:
                df_result = st.session_state["sg_upc_updated_df"]
                st.markdown(f"### 📋 成功新增 UPC 預覽（共 {len(df_result)} 筆）")
                st.dataframe(df_result, use_container_width=True)
                
                towrite_sg = io.BytesIO()
                with pd.ExcelWriter(towrite_sg, engine="openpyxl") as writer:
                    df_result.to_excel(writer, index=False)
                    
                download_filename = f"batch_edit_upc_added_only_{datetime.date.today().strftime('%Y%m%d')}.xlsx"
                
                st.download_button(
                    label=f"📥 下載包含 {len(df_result)} 筆成功自動填補的 Sitegiant 檔案",
                    data=towrite_sg.getvalue(),
                    file_name=download_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )

    # -------------------------------------------------------------------------
    # 子功能 4：📋 採購單待處理
    # -------------------------------------------------------------------------
    elif sub_page == "📋 採購單待處理":
        st.subheader("📋 採購單待處理 (尚未建立商品清單)")
        TARGET_SHEET_ID = "1Ixp9V_u2yU8hiWhxQCHNDB4kPKlxDGD2"
        
        with st.spinner("⏳ 正在由雲端獲取待處理清單..."):
            try:
                file_bytes = get_cached_gdrive_file_bytes(TARGET_SHEET_ID)
                engine_kw = {"engine": "calamine"} if HAS_CALAMINE else {}
                df_pending = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, dtype=str, **engine_kw)
                
                df_pending.columns = df_pending.columns.astype(str).str.strip()
                
                if df_pending.empty:
                    st.success("🎉 太棒了！目前沒有任何待處理的異常商品與採購單。")
                else:
                    st.markdown(f"📊 **目前待處理總筆數**：`{len(df_pending)} 筆`")
                    st.info("💡 下方為過去轉換入庫單時，找不到雲端統整表對應紀錄的異常商品，請調閱並盡速前往建立或更新資料。")
                    
                    st.dataframe(df_pending, use_container_width=True)
                    
                    towrite_pending = io.BytesIO()
                    with pd.ExcelWriter(towrite_pending, engine='openpyxl') as writer:
                        df_pending.to_excel(writer, index=False, sheet_name="尚未建立商品清單")
                    
                    st.download_button(
                        label="📥 下載待處理清單報表 (.xlsx)",
                        data=towrite_pending.getvalue(),
                        file_name=f"尚未建立商品清單_匯出_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
                    if st.button("🔄 重新載入最新雲端資料", use_container_width=True):
                        get_cached_gdrive_file_bytes.clear()
                        st.rerun()
                        
            except Exception as e:
                st.error(f"❌ 讀取雲端清單失敗，可能是該檔案尚未被系統自動建立或權限不足。")

    elif sub_page == "🗓️ 預購追蹤":
        st.caption("從廠商單帶入款項、補 SKU 後對訂單、結單下載兩檔；到貨後再催款。寫回只覆寫既有雲端檔。")
        reload = st.button("重新載入預購追蹤", use_container_width=True, key="preorder_reload")
        if reload or "preorder_loaded" not in st.session_state:
            if reload:
                get_cached_gdrive_file_bytes.clear()
                _clear_preorder_campaign_editor()
                st.session_state.pop("preorder_orders_loaded", None)
                st.session_state.pop("preorder_lock_result", None)
                st.session_state.pop("preorder_lock_preview", None)
                st.session_state.pop("preorder_vendor_history_before_lock", None)
                st.session_state.pop("preorder_sku_status_editor", None)
                st.session_state.pop("preorder_sku_editor_skus", None)
                st.session_state.pop("preorder_pending_source", None)
                st.session_state.pop("preorder_orders_synced", None)
            loaded = load_preorder_tracker()
            st.session_state["preorder_loaded"] = loaded
            if loaded.get("ok"):
                st.session_state["preorder_campaign_df"] = loaded["campaign"]
                st.session_state["preorder_vendor_history_df"] = loaded["vendor_history"]
                st.session_state["preorder_other_sheets"] = loaded.get("other_sheets") or {}
                st.session_state["preorder_pending_df"] = loaded.get("pending")
                st.session_state["preorder_sku_status_df"] = loaded.get("sku_status")
                st.session_state["preorder_notify_df"] = loaded.get("notify")
                st.session_state["preorder_tracker_name"] = loaded.get("name") or PREORDER_TRACKER_NAME

        loaded = st.session_state.get("preorder_loaded") or {}
        if not loaded.get("ok"):
            st.error(f"❌ 無法載入預購追蹤：{loaded.get('reason') or '未知錯誤'}")
            st.info("請確認 service account 對該檔有編輯權，且檔案是 .xlsx。")
            return

        if st.session_state.pop("preorder_save_ok", False):
            st.success("已覆寫雲端預購追蹤.xlsx（未新建檔、未改 SiteGiant 採購單空殼）。")
        import_msg = st.session_state.pop("preorder_vendor_import_msg", None)
        if import_msg:
            st.success(import_msg)

        st.caption(
            f"雲端檔：`{loaded.get('name')}`　最後修改：`{format_gdrive_time(loaded.get('modified'))}`"
        )

        with st.expander("計算規則"):
            st.markdown(
                "只計商品名稱含「預購」的列，再依活動 SKU 對 All Orders 庫存SKU（空白才改用商品SKU）。"
                "已取消不算進已接單。付款只認已付款／未付款（已退款不催款、不可打單）。\n\n"
                "鎖定數量 = 自購 + min(客戶量, 上限)。匯入不覆蓋 SKU／自購／上限。"
                "催款只含已標到貨的 SKU。官方訂單狀態見 "
                "[SiteGiant 說明](https://support.sitegiant.com/zh/knowledge-base/how-the-order-payment-and-fulfillment-status-works-zh/#2-e8a882e596aee78b80e6858b)。"
            )

        restock = probe_sg_restock_template()
        if not restock.get("ok"):
            st.error(f"❌ 無法讀取 SiteGiant 採購單空殼：{restock.get('reason') or '未知錯誤'}")
        with st.expander("SiteGiant 採購單空殼（進階）"):
            if restock.get("ok"):
                st.caption(
                    f"空殼可下載：`{restock.get('name')}`"
                    f"（{restock.get('nbytes', 0)} bytes，未寫回此檔）。結單時請下載已填數量的檔。"
                )
                st.download_button(
                    label="本機下載官方空殼（不會改雲端）",
                    data=restock.get("bytes") or b"",
                    file_name=restock.get("name") or "import_restock.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="preorder_restock_shell_dl",
                )
            else:
                st.caption("空殼連線失敗時，無法產出結單用的 SiteGiant 採購單。")

        edited = st.session_state.get("preorder_campaign_df", loaded["campaign"])
        edited = _render_preorder_vendor_import(edited)
        st.session_state["preorder_campaign_df"] = edited

        st.markdown("### 2. 補賣場資料並存檔")
        st.caption("SKU 填自定義編碼（與 SiteGiant 庫存 SKU 同一顆）。貨號來自廠商單，不要填進 SKU。")
        edited = st.data_editor(
            st.session_state.get("preorder_campaign_df", loaded["campaign"]),
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_order=list(PREORDER_CAMPAIGN_UI_COLUMN_ORDER),
            column_config={
                "月份": st.column_config.TextColumn("月份", help="YYYY-MM"),
                "結單日": st.column_config.TextColumn("結單日", help="YYYY-MM-DD"),
                "SKU": st.column_config.TextColumn("SKU", help="自定義編碼／庫存 SKU；四款請填四個不同 SKU。匯入廠商單不會填這欄。"),
                "條碼": st.column_config.TextColumn("條碼", help="GTIN／c"),
                "貨號": st.column_config.TextColumn("貨號", help="廠商貨號（來自訂購單，不是庫存 SKU）"),
                "品名": st.column_config.TextColumn("品名"),
                "廠商檔名": st.column_config.TextColumn("廠商檔名", help="來源訂購單檔名"),
                "自購": st.column_config.NumberColumn("自購", min_value=0, step=1),
                "上限": st.column_config.NumberColumn("上限", help="客戶預購上限", min_value=0, step=1),
                "私密連結": st.column_config.TextColumn("私密連結"),
                "預計關閉": st.column_config.TextColumn("預計關閉", help="預計結單日"),
                "實際關閉": st.column_config.TextColumn("實際關閉", help="實際關掉接受缺貨的時間"),
            },
            key="preorder_campaign_editor_v4",
        )
        st.session_state["preorder_campaign_df"] = edited
        dups = duplicate_preorder_skus(edited)
        if dups:
            st.warning("同一 SKU 出現在多列，看板客戶量會重複加總：" + "、".join(f"`{s}`" for s in dups))
        sku_as_item = preorder_rows_sku_equals_item_no(edited)
        if sku_as_item:
            st.warning(
                "有列的 SKU 與廠商貨號相同（"
                + "、".join(f"`{s}`" for s in sku_as_item)
                + "）。庫存 SKU 請填 SiteGiant／蝦皮自定義編碼，貨號請留在「貨號」欄。"
            )
        empty_skus = _empty_preorder_sku_labels(edited)
        if empty_skus:
            st.warning("待補 SKU：" + "、".join(empty_skus[:12]) + ("…" if len(empty_skus) > 12 else ""))

        locked = bool(st.session_state.get("preorder_lock_result"))
        save_col, dl_col = st.columns(2)
        with save_col:
            _save_preorder_tracker_button(edited, "preorder_save_campaign", primary=not locked)
        with dl_col:
            local_bytes = preorder_tracker_bytes(
                edited,
                st.session_state.get("preorder_vendor_history_df"),
                st.session_state.get("preorder_other_sheets"),
                pending_df=st.session_state.get("preorder_pending_df"),
                sku_status_df=st.session_state.get("preorder_sku_status_df"),
                notify_df=st.session_state.get("preorder_notify_df"),
            )
            st.download_button(
                label="下載目前活動表（本機）",
                data=local_bytes,
                file_name=st.session_state.get("preorder_tracker_name") or PREORDER_TRACKER_NAME,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="preorder_tracker_local_dl",
            )

        st.write("---")
        orders_df = _render_preorder_orders_board(edited)
        st.write("---")
        _render_preorder_close(edited, orders_df)
        st.write("---")
        _render_preorder_arrival_copy(
            edited,
            orders_df,
            st.session_state.get("preorder_sku_status_df"),
        )
