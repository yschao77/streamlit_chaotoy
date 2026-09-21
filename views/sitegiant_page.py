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
    preorder_lock_selection_df,
    append_preorder_vendor_history,
    fill_vendor_po_qty_bytes,
    sync_preorder_vendor_lock_record,
    fill_sg_restock_bytes,
    taipei_now,
    PREORDER_VENDOR_HISTORY_COLUMNS,
    PREORDER_STATUS_PREORDER,
    PREORDER_STATUS_ARRIVED,
    extract_preorder_pending_df,
    attach_preorder_notify_days,
    upsert_preorder_sku_status,
    stamp_preorder_first_notify,
    preorder_notify_keys_from_lines,
    build_preorder_notify_display,
    campaign_df_from_arrived_skus,
    arrived_skus_missing_from_campaign,
    ensure_preorder_sku_status_df,
    compute_preorder_campaign_stages,
    unpaid_skus_from_board,
    sku_accepted_qty_map,
    PREORDER_STAGE_ORDER,
    PREORDER_STAGE_DONE,
    PREORDER_STAGE_AWAITING_ARRIVAL,
    ensure_preorder_campaign_df,
    _preorder_self_buy,
    PREORDER_CLOSED_LOCKED_COLUMNS,
    validate_preorder_campaign_cutoffs,
    build_preorder_sku_overview,
    apply_sku_overview_status_edits,
    preorder_sku_detail_lines,
    PREORDER_SKU_OVERVIEW_COLUMNS,
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
    "賣場後台連結",
    "預計關閉",
    "實際關閉",
    "日曆事件ID",
)
PREORDER_LOCK_DISPLAY_COLUMNS = ("SKU", "條碼", "貨號", "自購", "客戶量", "上限", "鎖定數量")


def _editor_base(base_key, df, *, replace_empty=False):
    base = st.session_state.get(base_key)
    if base is None:
        st.session_state[base_key] = df.copy() if df is not None else df
    elif (
        replace_empty
        and getattr(base, "empty", True)
        and df is not None
        and not getattr(df, "empty", True)
    ):
        st.session_state[base_key] = df.copy()
    return st.session_state[base_key]


def _clear_preorder_campaign_editor():
    st.session_state.pop("preorder_campaign_editor_open_v1", None)
    st.session_state.pop("preorder_campaign_editor_closed_v1", None)
    st.session_state.pop("preorder_editor_open_base", None)
    st.session_state.pop("preorder_editor_closed_base", None)
    st.session_state.pop("preorder_stage_move_msg", None)
    # 舊版單一表格的 key／base，重複 pop 是安全的（本來就不存在時 pop 不會出錯）。
    st.session_state.pop("preorder_campaign_editor_v5", None)
    st.session_state.pop("preorder_campaign_editor_v4", None)
    st.session_state.pop("preorder_campaign_editor_v3", None)
    st.session_state.pop("preorder_campaign_editor_v2", None)
    st.session_state.pop("preorder_campaign_editor", None)
    st.session_state.pop("preorder_editor_base", None)


def _clear_preorder_lock_select_editor():
    st.session_state.pop("preorder_lock_select_editor", None)
    st.session_state.pop("preorder_lock_select_base", None)


def _clear_inward_grid_editor():
    st.session_state.pop("inward_grid", None)
    st.session_state.pop("inward_grid_base", None)


def _clear_preorder_sku_overview_state():
    """重新載入雲端資料後，清掉「3. 對客戶訂單」彙總表/搜尋框/明細選取，
    以及「5. 到貨催款」標到貨編輯表的殘留 widget 狀態。"""
    st.session_state.pop("preorder_sku_overview_editor", None)
    st.session_state.pop("preorder_sku_overview_search", None)
    st.session_state.pop("preorder_sku_detail_pick", None)
    st.session_state.pop("preorder_pending_search", None)
    st.session_state.pop("preorder_arrival_status_editor", None)
    st.session_state.pop("preorder_arrival_status_search", None)
    # 舊版獨立 SKU 狀態編輯表的 key，重複 pop 是安全的。
    st.session_state.pop("preorder_sku_status_editor", None)
    st.session_state.pop("preorder_sku_editor_base", None)
    st.session_state.pop("preorder_sku_editor_skus", None)


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
    result = _persist_preorder_tracker(campaign_df)
    if result.get("ok"):
        st.session_state["preorder_save_ok"] = True
        get_cached_gdrive_file_bytes.clear()
        st.session_state.pop("preorder_loaded", None)
        st.rerun()
    st.error(f"❌ 寫回失敗：{result.get('reason') or '未知錯誤'}")


def _persist_preorder_tracker(campaign_df):
    cutoff_errors = validate_preorder_campaign_cutoffs(campaign_df)
    if cutoff_errors:
        shown = cutoff_errors[:12]
        reason = "結單日驗證未通過（有 SKU 的列必填，格式 YYYY-MM-DD HH:mm）：\n" + "\n".join(shown)
        if len(cutoff_errors) > len(shown):
            reason += f"\n…另有 {len(cutoff_errors) - len(shown)} 筆"
        return {"ok": False, "reason": reason}
    return save_preorder_tracker(
        campaign_df,
        st.session_state.get("preorder_vendor_history_df"),
        st.session_state.get("preorder_other_sheets"),
        file_name=st.session_state.get("preorder_tracker_name"),
        pending_df=st.session_state.get("preorder_pending_df"),
        sku_status_df=st.session_state.get("preorder_sku_status_df"),
        notify_df=st.session_state.get("preorder_notify_df"),
    )


def _preorder_fragment(fn):
    deco = getattr(st, "fragment", None) or getattr(st, "experimental_fragment", None)
    return deco(fn) if deco else fn


PREORDER_STAGE_ICONS = {
    "收單中": "🟡",
    "已結單待到貨": "🟠",
    "催款中": "🔴",
    "完成": "🟢",
}


def _render_preorder_lifecycle_overview(
    campaign_df, sku_status_df, unpaid_skus=None, accepted_qty_by_sku=None
):
    """頁面最上方的生命週期總覽：依「實際關閉」＋「SKU狀態」把每個活動列分到 4 個階段。"""
    staged = compute_preorder_campaign_stages(campaign_df, sku_status_df, unpaid_skus)
    if staged.empty:
        st.caption("活動表沒有列，暫無生命週期總覽。")
        return

    if accepted_qty_by_sku is not None:
        no_order_skus = {sku for sku, qty in accepted_qty_by_sku.items() if not qty}
        drop = (staged["階段"] == PREORDER_STAGE_AWAITING_ARRIVAL) & staged["SKU"].map(
            _preorder_text
        ).isin(no_order_skus)
        staged = staged.loc[~drop]

    counts = staged["階段"].value_counts()
    cols = st.columns(len(PREORDER_STAGE_ORDER))
    for col, stage in zip(cols, PREORDER_STAGE_ORDER):
        icon = PREORDER_STAGE_ICONS.get(stage, "")
        col.metric(f"{icon} {stage}", int(counts.get(stage, 0)))

    captions = []
    if unpaid_skus is None:
        captions.append(
            "「催款中」目前只要已到貨就列在這裡，還沒對到 All Orders 未付款件數，"
            "所以看不到「完成」；往下拉到「3. 對客戶訂單」載入 All Orders 後這裡會自動補上。"
        )
    if accepted_qty_by_sku is None:
        captions.append(
            "「已結單待到貨」目前還沒排除沒有預購訂單的 SKU；"
            "往下拉到「3. 對客戶訂單」載入 All Orders 後這裡會自動篩掉。"
        )
    if captions:
        st.caption(" ".join(captions))

    show_cols = [c for c in ("月份", "SKU", "品名", "結單日", "實際關閉") if c in staged.columns]
    for stage in PREORDER_STAGE_ORDER:
        rows = staged.loc[staged["階段"] == stage]
        icon = PREORDER_STAGE_ICONS.get(stage, "")
        with st.expander(f"{icon} {stage}（{len(rows)}）", expanded=False):
            if rows.empty:
                st.caption("目前沒有列在這個階段。")
            else:
                st.dataframe(rows[show_cols], use_container_width=True, hide_index=True)


def _preorder_campaign_column_config(disabled_cols=()):
    d = lambda col: col in disabled_cols  # noqa: E731
    return {
        "月份": st.column_config.TextColumn("月份", help="YYYY-MM", disabled=d("月份")),
        "結單日": st.column_config.TextColumn(
            "結單日",
            help="YYYY-MM-DD HH:mm（有 SKU 必填；僅第 2 段可編）",
            disabled=d("結單日"),
        ),
        "SKU": st.column_config.TextColumn(
            "SKU",
            help="自定義編碼／庫存 SKU；四款請填四個不同 SKU。匯入廠商單不會填這欄。",
            disabled=d("SKU"),
        ),
        "條碼": st.column_config.TextColumn("條碼", help="GTIN／c"),
        "貨號": st.column_config.TextColumn("貨號", help="廠商貨號（來自訂購單，不是庫存 SKU）"),
        "品名": st.column_config.TextColumn("品名"),
        "廠商檔名": st.column_config.TextColumn("廠商檔名", help="來源訂購單檔名"),
        "自購": st.column_config.NumberColumn(
            "自購", min_value=0, step=1, help="已實際關閉的列不能改", disabled=d("自購")
        ),
        "上限": st.column_config.NumberColumn(
            "上限", help="客戶預購上限。已實際關閉的列不能改。", min_value=0, step=1, disabled=d("上限")
        ),
        "私密連結": st.column_config.TextColumn("私密連結"),
        "賣場後台連結": st.column_config.TextColumn(
            "賣場後台連結", help="寫入日曆事件描述；進行中／已關閉皆可改"
        ),
        "預計關閉": st.column_config.TextColumn("預計關閉", help="預計結單日"),
        "實際關閉": st.column_config.TextColumn(
            "實際關閉", help="有值＝已關閉。清空即可重開，重開後這列會移回上方可編輯表。"
        ),
        "日曆事件ID": st.column_config.TextColumn(
            "日曆事件ID", help="系統維護，請勿手改", disabled=True
        ),
    }


def _preorder_closed_flags(df):
    if df is None or getattr(df, "empty", True):
        return pd.Series(dtype=bool)
    return df.get("實際關閉", pd.Series([""] * len(df), index=df.index)).map(_preorder_text).astype(bool)


@_preorder_fragment
def _render_preorder_campaign_editor():
    loaded = st.session_state.get("preorder_loaded") or {}
    campaign = ensure_preorder_campaign_df(
        st.session_state.get("preorder_campaign_df", loaded.get("campaign"))
    )
    closed_flags = _preorder_closed_flags(campaign)
    open_rows = campaign.loc[~closed_flags].reset_index(drop=True)
    closed_rows = campaign.loc[closed_flags].reset_index(drop=True)

    st.caption(
        "SKU 填自定義編碼（與 SiteGiant 庫存 SKU 同一顆）。貨號來自廠商單，不要填進 SKU。"
        " 「進行中」列可自由編輯；填了實際關閉會立刻移到下方「已關閉」收合區，"
        "自購／上限／SKU／月份／結單日在那裡直接鎖住不能改（不必送出才還原）。"
        " 品名、私密連結、賣場後台連結、條碼、貨號兩邊都能改；清空實際關閉即可重開並移回這裡。"
        " 結單日格式 YYYY-MM-DD HH:mm（有 SKU 必填）；日曆事件ID 系統維護請勿手改。"
    )

    open_base = _editor_base("preorder_editor_open_base", open_rows, replace_empty=True)
    edited_open = st.data_editor(
        open_base,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_order=list(PREORDER_CAMPAIGN_UI_COLUMN_ORDER),
        column_config=_preorder_campaign_column_config(),
        key="preorder_campaign_editor_open_v1",
    )
    edited_open = ensure_preorder_campaign_df(edited_open)

    with st.expander(
        f"🔒 已關閉（{len(closed_rows)} 列，自購／上限／SKU／月份／結單日已鎖定）",
        expanded=False,
    ):
        if closed_rows.empty:
            st.caption("目前沒有已關閉的活動列。")
            edited_closed = closed_rows
        else:
            st.caption(
                "這些列已結單，鎖定欄位直接鎖住不給打字。品名／私密連結／賣場後台連結／條碼／貨號仍可改。"
                "清空實際關閉會移回上方「進行中」表；如需整列刪除，也請先清空再到上方刪。"
            )
            closed_base = _editor_base("preorder_editor_closed_base", closed_rows, replace_empty=True)
            edited_closed = st.data_editor(
                closed_base,
                num_rows="fixed",
                use_container_width=True,
                hide_index=True,
                column_order=list(PREORDER_CAMPAIGN_UI_COLUMN_ORDER),
                column_config=_preorder_campaign_column_config(PREORDER_CLOSED_LOCKED_COLUMNS),
                key="preorder_campaign_editor_closed_v1",
            )
            edited_closed = ensure_preorder_campaign_df(edited_closed)

    open_became_closed = int(_preorder_closed_flags(edited_open).sum())
    closed_became_open = int((~_preorder_closed_flags(edited_closed)).sum()) if len(edited_closed) else 0

    merged = ensure_preorder_campaign_df(pd.concat([edited_open, edited_closed], ignore_index=True))
    st.session_state["preorder_campaign_df"] = merged

    if open_became_closed or closed_became_open:
        st.session_state.pop("preorder_campaign_editor_open_v1", None)
        st.session_state.pop("preorder_campaign_editor_closed_v1", None)
        st.session_state.pop("preorder_editor_open_base", None)
        st.session_state.pop("preorder_editor_closed_base", None)
        bits = []
        if open_became_closed:
            bits.append(f"{open_became_closed} 列剛填了實際關閉，已移到下方「已關閉」收合區。")
        if closed_became_open:
            bits.append(f"{closed_became_open} 列清空了實際關閉，已移回上方「進行中」表。")
        st.session_state["preorder_stage_move_msg"] = "".join(bits)
        try:
            st.rerun(scope="fragment")
        except TypeError:
            st.rerun()
        return

    move_msg = st.session_state.pop("preorder_stage_move_msg", None)
    if move_msg:
        st.info(move_msg)
    dups = duplicate_preorder_skus(merged)
    if dups:
        st.warning("同一 SKU 出現在多列，看板客戶量會重複加總：" + "、".join(f"`{s}`" for s in dups))
    sku_as_item = preorder_rows_sku_equals_item_no(merged)
    if sku_as_item:
        st.warning(
            "有列的 SKU 與廠商貨號相同（"
            + "、".join(f"`{s}`" for s in sku_as_item)
            + "）。庫存 SKU 請填 SiteGiant／蝦皮自定義編碼，貨號請留在「貨號」欄。"
        )
    empty_skus = _empty_preorder_sku_labels(merged)
    if empty_skus:
        st.warning("待補 SKU：" + "、".join(empty_skus[:12]) + ("…" if len(empty_skus) > 12 else ""))


def _show_preorder_lock_tables(result):
    skipped = result.get("skipped") or []
    locked_rows = result.get("locked_rows") or []
    if locked_rows:
        show = pd.DataFrame(locked_rows)
        cols = [c for c in PREORDER_LOCK_DISPLAY_COLUMNS if c in show.columns]
        st.markdown(f"**要鎖定 {len(show)} 列**")
        st.dataframe(show[cols], hide_index=True, use_container_width=True)
    else:
        st.info("這次沒有要鎖定的列。")
    if skipped:
        closed = [s for s in skipped if "實際關閉" in str(s.get("原因") or "")]
        unchecked = [s for s in skipped if "未勾選" in str(s.get("原因") or "")]
        no_sku = [s for s in skipped if s not in closed and s not in unchecked]
        with st.expander(f"不鎖定／跳過的列（{len(skipped)} 列）"):
            if no_sku:
                st.warning(f"{len(no_sku)} 列因未填 SKU 未鎖定。")
                st.dataframe(pd.DataFrame(no_sku), hide_index=True, use_container_width=True)
            if closed:
                st.info(f"{len(closed)} 列已實際關閉，不重算訂量。清空實際關閉後才可再鎖定。")
                st.dataframe(pd.DataFrame(closed), hide_index=True, use_container_width=True)
            if unchecked:
                st.info(f"{len(unchecked)} 列未勾選結單，這次不鎖定。")
                st.dataframe(pd.DataFrame(unchecked), hide_index=True, use_container_width=True)


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
    st.session_state.pop("preorder_lock_preview_self_buy", None)
    st.session_state.pop("preorder_lock_selected", None)
    _clear_preorder_lock_select_editor()


def _render_preorder_orders_board(campaign_df):
    st.markdown("### 3. 對客戶訂單")
    st.caption("摘要看板與 SKU 狀態。本機 Orders 只預覽、不會上傳到 Drive。")

    reload_orders = st.button(
        "🔄 重新載入雲端 All Orders",
        use_container_width=True,
        key="preorder_orders_reload",
    )
    st.caption("只重抓 All Orders 與訂單看板；不影響上方活動表尚未存檔的編輯，也不會清掉廠商檔快取或結單預覽。")
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
            type=["xlsx", "xls", "zip"],
            key="preorder_orders_local",
            help="可上傳 Orders_DD-MM-YYYY-*.zip 或已解開的 xlsx。不會 files.create／update 到雲端。",
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
    drive_inner = drive_loaded.get("inner_name")
    inner_note = f"　解出：`{drive_inner}`" if drive_inner and drive_inner != drive_name else ""
    st.caption(f"雲端最新 All Orders：`{drive_name}`　最後修改：`{drive_time}`{inner_note}")

    if loaded_orders.get("preview"):
        local_inner = loaded_orders.get("inner_name")
        unzip_note = f"（解出 `{local_inner}`）" if local_inner and local_inner != loaded_orders.get("name") else ""
        st.success(f"目前用本機檔預覽：`{loaded_orders.get('name')}`{unzip_note}（未寫入 Drive）")

    if not loaded_orders.get("ok"):
        st.error(f"❌ 無法載入 All Orders：{loaded_orders.get('reason') or '未知錯誤'}")
        st.info("可改本機上傳 zip／xlsx 預覽。請確認 service account 對 All Orders 資料夾有檢視權，且檔名是 Orders_DD-MM-YYYY-*.zip（內含 xlsx）。")
        st.session_state["preorder_orders_synced"] = False
        sku_status = ensure_preorder_sku_status_df(st.session_state.get("preorder_sku_status_df"))
        pending = st.session_state.get("preorder_pending_df")
        sku_status = _render_preorder_sku_overview_and_detail(
            None, sku_status, pending, campaign_df, editable=False
        )
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
    st.session_state["preorder_pending_df"] = pending
    st.session_state["preorder_sku_status_df"] = sku_status
    st.session_state["preorder_orders_synced"] = True

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
        st.info("活動表沒有列，看板還沒有可對照的活動；下面仍會列出 All Orders 出現過的 SKU。")

    sku_status = _render_preorder_sku_overview_and_detail(
        board, sku_status, pending, campaign_df, editable=False
    )
    st.session_state["preorder_sku_status_df"] = sku_status
    return orders_df


def _filter_preorder_overview(df, keyword):
    keyword = (keyword or "").strip().lower()
    if not keyword or df is None or df.empty:
        return df
    mask = (
        df["SKU"].map(_preorder_text).str.lower().str.contains(keyword, regex=False)
        | df["品名"].map(_preorder_text).str.lower().str.contains(keyword, regex=False)
    )
    return df.loc[mask].reset_index(drop=True)


def _filter_preorder_pending(df, keyword):
    keyword = (keyword or "").strip().lower()
    if not keyword or df is None or getattr(df, "empty", True):
        return df
    cols = [c for c in ("SKU", "訂單編號", "商品名稱") if c in df.columns]
    if not cols:
        return df
    mask = None
    for c in cols:
        col_mask = df[c].map(_preorder_text).str.lower().str.contains(keyword, regex=False)
        mask = col_mask if mask is None else (mask | col_mask)
    return df.loc[mask].reset_index(drop=True)


def _render_preorder_sku_overview_editor(overview_df, *, key="preorder_sku_overview_editor"):
    # 只有「狀態」這一欄可編輯，且是單選 selectbox（原子提交，不是連續打字），跟現有
    # 勾選 checkbox 表（廠商匯入、結單候選）風險屬性一樣，所以不用 `_editor_base` 凍結
    # 底稿那套（那套是給自購/上限這類自由輸入欄位防「連改兩格跳回上一筆」用的）。這裡
    # 的 data 本來就要隨搜尋框即時變動，凍結底稿反而會讓搜尋失效。
    edited = st.data_editor(
        overview_df,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_order=list(PREORDER_SKU_OVERVIEW_COLUMNS),
        column_config={
            "SKU": st.column_config.TextColumn("SKU", disabled=True),
            "品名": st.column_config.TextColumn("品名", disabled=True),
            "階段": st.column_config.TextColumn("階段", disabled=True),
            "已接單": st.column_config.NumberColumn("已接單", disabled=True),
            "上限": st.column_config.TextColumn("上限", disabled=True),
            "已付款件數": st.column_config.NumberColumn("已付款件數", disabled=True),
            "未付款件數": st.column_config.NumberColumn("未付款件數", disabled=True),
            "達上限": st.column_config.TextColumn("達上限", disabled=True),
            "狀態": st.column_config.SelectboxColumn(
                "狀態",
                options=[PREORDER_STATUS_PREORDER, PREORDER_STATUS_ARRIVED],
                required=True,
            ),
            "在活動表": st.column_config.TextColumn("在活動表", disabled=True),
        },
        key=key,
    )
    return edited


def _preorder_stage_by_sku(board, campaign_df, sku_status):
    """活動 SKU → 階段；無訂單的「已結單待到貨」會改成空字串（與總覽／第3段口徑一致）。"""
    stage_unpaid = unpaid_skus_from_board(board) if board else None
    staged_campaign = compute_preorder_campaign_stages(campaign_df, sku_status, stage_unpaid)
    _accepted_qty_by_sku = sku_accepted_qty_map(board) if board else {}
    _no_order_skus = {sku for sku, qty in _accepted_qty_by_sku.items() if not qty}
    stage_by_sku = {}
    for _, r in staged_campaign.iterrows():
        r_sku = _preorder_text(r.get("SKU"))
        if r_sku:
            r_stage = r.get("階段")
            if r_stage == PREORDER_STAGE_AWAITING_ARRIVAL and r_sku in _no_order_skus:
                r_stage = ""
            stage_by_sku[r_sku] = r_stage
    return stage_by_sku


def _render_preorder_sku_overview_and_detail(board, sku_status, pending, campaign_df, editable=False):
    """「3. 對客戶訂單」：可搜尋的彙總表＋選取才渲染明細。狀態欄預設唯讀；
    要標到貨請到「5. 到貨催款」。"""
    sku_status = ensure_preorder_sku_status_df(sku_status)
    missing = arrived_skus_missing_from_campaign(sku_status, campaign_df)
    stage_by_sku = _preorder_stage_by_sku(board, campaign_df, sku_status)

    # 已達上限但這個 SKU 已經是「完成」（到貨且沒有未付款件數）就不用再提醒了——
    # 上限是給「還在收單/催款」時判斷要不要繼續接單用的，結案的批次不需要再跳警告。
    over_active_skus = []
    if board:
        for item in board.get("campaigns") or []:
            if not item.get("over_limit"):
                continue
            sku = _preorder_text(item.get("sku"))
            if stage_by_sku.get(sku) == PREORDER_STAGE_DONE:
                continue
            over_active_skus.append(sku or "（未填 SKU）")

    tab_board, tab_raw = st.tabs(["📋 看板與明細", "🧾 原始訂單明細"])

    with tab_board:
        if missing:
            st.warning(
                "到貨 SKU 不在活動表（不自動寫入；要改狀態請到「5. 到貨催款」）："
                + "、".join(f"`{s}`" for s in missing)
            )
        if over_active_skus:
            st.warning("⚠️ 已接單已達或超過上限：" + "、".join(f"`{sku}`" for sku in over_active_skus))

        overview = build_preorder_sku_overview(board, sku_status, stage_by_sku)
        if overview.empty:
            st.info("目前沒有任何 SKU（活動表沒填 SKU，也還沒有 All Orders 資料）。")
        else:
            search = st.text_input(
                "🔎 搜尋 SKU／品名",
                key="preorder_sku_overview_search",
                placeholder="輸入關鍵字篩選下面表格",
            )
            filtered = _filter_preorder_overview(overview, search)
            if editable:
                edited_overview = _render_preorder_sku_overview_editor(filtered)
                sku_status = apply_sku_overview_status_edits(sku_status, edited_overview)
                st.session_state["preorder_sku_status_df"] = sku_status
                st.caption("改「狀態」後還沒寫回雲端，按下面「覆寫雲端預購追蹤」才會存檔。")
                _save_preorder_tracker_button(campaign_df, "preorder_save_sku_overview", primary=False)
            else:
                edited_overview = filtered
                st.dataframe(filtered, use_container_width=True, hide_index=True)
            st.caption(
                "已付款／未付款依付款狀態，只顯示看板口徑；可打單另要求訂單狀態＝待處理。"
                "「狀態」欄在此唯讀——要標到貨請到「5. 到貨催款」。"
                "「在活動表」＝否表示這個 SKU 還沒建活動列，數字欄留空。"
            )

            sku_options = [s for s in edited_overview["SKU"].tolist() if s]
            picked = st.multiselect(
                "選要看明細的 SKU（已付款／未付款合併顯示，未付款排前面）",
                options=sku_options,
                key="preorder_sku_detail_pick",
            )
            if picked and board:
                detail = preorder_sku_detail_lines(board, picked)
                detail = attach_preorder_notify_days(detail, st.session_state.get("preorder_notify_df"))
                st.dataframe(detail, use_container_width=True, hide_index=True)
            elif picked:
                st.info("目前沒有訂單資料可顯示明細（All Orders 還沒載入）。")
            else:
                st.caption("在上面選取 SKU 才會顯示逐筆明細，避免 SKU 一多整頁被撐長。")

    with tab_raw:
        st.caption(
            "最新 All Orders 快照，含所有預購訂單列（不分活動/SKU 是否對得上）。"
            "已過天數只算未付款、未取消、未退款。除錯／稽核用，跟上面看板數字口徑可能不完全一樣。"
        )
        pending_n = 0 if pending is None or getattr(pending, "empty", True) else len(pending)
        dated_n = 0
        if pending is not None and not getattr(pending, "empty", True) and "首次通知日" in pending.columns:
            dated_n = int(pending["首次通知日"].map(_preorder_text).astype(bool).sum())
        st.caption(f"共 {pending_n} 列，{dated_n} 列已有首次通知日。")
        if pending is None or getattr(pending, "empty", True):
            st.info("目前沒有預購訂單列。")
        else:
            raw_search = st.text_input(
                "🔎 搜尋 SKU／訂單編號／商品名稱",
                key="preorder_pending_search",
                placeholder="輸入關鍵字篩選下面表格",
            )
            raw_filtered = _filter_preorder_pending(pending, raw_search)
            st.dataframe(raw_filtered, use_container_width=True, hide_index=True)

    return sku_status


def _render_preorder_arrival_status_editor(campaign_df, orders_df, sku_status_df):
    """第 5 段開頭：唯一可改「預購／到貨」的地方。"""
    st.markdown("**標到貨**")
    st.caption(
        "這裡是唯一可改「預購／到貨」的地方。優先處理「已結單待到貨」；"
        "改成到貨後，下方催款／可打單會立刻依目前畫面重算。"
        " 改完按「覆寫雲端預購追蹤」才寫回雲端（與第 2 段同一份檔）。"
    )
    sku_status = ensure_preorder_sku_status_df(sku_status_df)
    missing = arrived_skus_missing_from_campaign(sku_status, campaign_df)
    if missing:
        st.warning(
            "到貨 SKU 不在活動表（不自動寫入，仍可在下表改狀態）："
            + "、".join(f"`{s}`" for s in missing)
        )

    board = build_preorder_board(campaign_df, orders_df) if orders_df is not None else None
    stage_by_sku = _preorder_stage_by_sku(board, campaign_df, sku_status)
    overview = build_preorder_sku_overview(board, sku_status, stage_by_sku)
    if overview.empty:
        st.info("目前沒有可標到貨的 SKU（活動表沒填 SKU，也還沒有 All Orders／SKU 狀態）。")
        return sku_status

    # 只列尚未標到貨（狀態＝預購）；已到貨留在 SKU 狀態表，不出現在這張編輯表。
    pending_arrival = overview.loc[
        overview["狀態"].map(_preorder_text) == PREORDER_STATUS_PREORDER
    ].reset_index(drop=True)
    awaiting_n = (
        int((pending_arrival["階段"] == PREORDER_STAGE_AWAITING_ARRIVAL).sum())
        if "階段" in pending_arrival.columns and not pending_arrival.empty
        else 0
    )
    st.caption(
        f"只列尚未標到貨（狀態＝預購）。其中「已結單待到貨」{awaiting_n} 列。"
    )
    if pending_arrival.empty:
        st.info("沒有待標到貨的 SKU（狀態都已是到貨，或尚無 SKU）。")
        return sku_status

    search = st.text_input(
        "🔎 搜尋要標到貨的 SKU／品名",
        key="preorder_arrival_status_search",
        placeholder="只搜尋狀態＝預購的列",
    )
    filtered = _filter_preorder_overview(pending_arrival, search)
    if filtered.empty:
        st.info("搜尋沒有符合的待標到貨 SKU。")
        return sku_status
    edited_overview = _render_preorder_sku_overview_editor(
        filtered, key="preorder_arrival_status_editor"
    )
    sku_status = apply_sku_overview_status_edits(sku_status, edited_overview)
    st.session_state["preorder_sku_status_df"] = sku_status
    _save_preorder_tracker_button(campaign_df, "preorder_save_arrival_status", primary=False)
    return sku_status


def _stamp_and_persist_notify(campaign_df, keys):
    result = stamp_preorder_first_notify(st.session_state.get("preorder_notify_df"), keys)
    st.session_state["preorder_notify_df"] = result["notify"]
    pending = st.session_state.get("preorder_pending_df")
    st.session_state["preorder_pending_df"] = attach_preorder_notify_days(
        pending, result["notify"]
    )
    save_result = _persist_preorder_tracker(campaign_df)
    if result.get("added"):
        msg = f"已為 {result['added']} 列點上首次通知日 {result.get('today')}（已有日期不改）。"
    else:
        msg = "沒有新的未付款列可點日期（可能已點過，或沒有到貨未付款單）。"
    if save_result.get("ok"):
        msg += " 已寫入雲端工作表「通知紀錄」與「預購訂單」。"
        get_cached_gdrive_file_bytes.clear()
        st.session_state.pop("preorder_loaded", None)
    else:
        msg += " 畫面已有日期，但雲端尚未寫回：" + (save_result.get("reason") or "未知錯誤")
    st.session_state["preorder_notify_msg"] = msg
    st.rerun()


def _render_preorder_arrival_copy(campaign_df, orders_df, sku_status_df):
    st.markdown("### 5. 到貨催款")
    st.caption(
        "先在本段「標到貨」改狀態；催款／可打單只含已標到貨的 SKU。"
        " 催款文＝未付款且非已取消／已退款。"
        " 可打單＝已付款且訂單狀態＝待處理。"
        " 按「本次催款已通知且覆寫雲端預購追蹤」會立刻寫入首次通知日，並覆寫雲端整份預購追蹤（含通知紀錄／預購訂單）。"
        " 打單、取消、勾已收到付款仍在 SiteGiant。"
    )
    sku_status_df = _render_preorder_arrival_status_editor(
        campaign_df, orders_df, sku_status_df
    )
    st.write("---")
    if orders_df is None:
        st.info("請先在「3. 對客戶訂單」載入 All Orders，才會出現催款／可打單名單。")
        return
    arrived_campaign = campaign_df_from_arrived_skus(sku_status_df)
    board = build_preorder_board(arrived_campaign, orders_df)
    copies = build_preorder_arrival_copy(board)
    paid_text = copies.get("paid_pick") or ""

    # 只在內容真的變動時才重建文字框元件，避免每次 rerun 都用 hash 換 key（游標/選取會被重置）。
    if st.session_state.get("preorder_copy_paid_src") != paid_text:
        st.session_state.pop("preorder_copy_paid", None)
        st.session_state["preorder_copy_paid_src"] = paid_text

    with st.expander(
        f"可打單名單（待處理＋已付款，{copies.get('paid_n', 0)} 列）", expanded=False
    ):
        st.text_area(
            "可複貼文字",
            value=paid_text,
            height=160,
            key="preorder_copy_paid",
            label_visibility="collapsed",
        )

    unpaid_lines_df = copies.get("unpaid_lines")
    if unpaid_lines_df is not None and not getattr(unpaid_lines_df, "empty", True):
        unpaid_view = attach_preorder_notify_days(
            unpaid_lines_df, st.session_state.get("preorder_notify_df")
        )
        with st.expander(f"逐筆選要通知的未付款訂單（{len(unpaid_view)} 列）"):
            st.caption(
                "已有首次通知日的列預設不勾，避免重複通知；整批蓋章也不會改既有日期。"
                " 全選可刻意全勾（重點名仍不覆寫日期）。"
            )
            display_cols = [
                c
                for c in [
                    "訂單編號",
                    "SKU",
                    "商品名稱",
                    "商品數量",
                    "顧客",
                    "金額",
                    "首次通知日",
                    "已過天數",
                ]
                if c in unpaid_view.columns
            ]
            select_base = unpaid_view[display_cols].reset_index(drop=True).copy()
            if "首次通知日" in select_base.columns:
                select_base.insert(
                    0,
                    "勾選",
                    ~select_base["首次通知日"].map(_preorder_text).astype(bool),
                )
            else:
                select_base.insert(0, "勾選", True)
            nsel_col1, nsel_col2 = st.columns(2)
            with nsel_col1:
                if st.button("全選", key="preorder_notify_select_all", use_container_width=True):
                    st.session_state.pop("preorder_notify_select_editor", None)
                    st.session_state["preorder_notify_select_override"] = True
                    st.rerun()
            with nsel_col2:
                if st.button("取消全選", key="preorder_notify_select_none", use_container_width=True):
                    st.session_state.pop("preorder_notify_select_editor", None)
                    st.session_state["preorder_notify_select_override"] = False
                    st.rerun()
            notify_override = st.session_state.pop("preorder_notify_select_override", None)
            if notify_override is not None:
                select_base["勾選"] = notify_override
            edited_select = st.data_editor(
                select_base,
                hide_index=True,
                use_container_width=True,
                disabled=display_cols,
                column_config={"勾選": st.column_config.CheckboxColumn("勾選", default=False)},
                key="preorder_notify_select_editor",
            )
            picked_rows = [
                i for i, rec in enumerate(edited_select.to_dict(orient="records")) if rec.get("勾選")
            ]
            if st.button(
                f"📌 只標記勾選列已通知且覆寫雲端預購追蹤（{len(picked_rows)} 列）",
                use_container_width=True,
                key="preorder_stamp_notify_selected",
            ):
                picked_keys = preorder_notify_keys_from_lines(unpaid_view.iloc[picked_rows])
                if not picked_keys:
                    st.warning("沒有勾選任何列。")
                else:
                    _stamp_and_persist_notify(campaign_df, picked_keys)

    if st.button(
        "📌 本次催款已通知且覆寫雲端預購追蹤（整批）",
        use_container_width=True,
        key="preorder_stamp_notify",
    ):
        keys = preorder_notify_keys_from_lines(copies.get("unpaid_lines"))
        _stamp_and_persist_notify(campaign_df, keys)
    notify_msg = st.session_state.pop("preorder_notify_msg", None)
    if notify_msg:
        st.success(notify_msg)
    notify = st.session_state.get("preorder_notify_df")
    if notify is not None and not getattr(notify, "empty", True):
        st.markdown("**通知紀錄**")
        notify_view = build_preorder_notify_display(notify, orders_df)
        st.caption(
            "只列仍未付款（已付款／已退款／已取消不顯示）。"
            "已過天數供判斷是否取消；本頁不取消。工作表仍保留全部通知列。"
        )
        if notify_view is None or getattr(notify_view, "empty", True):
            st.info("目前沒有仍未付款的通知紀錄（可能都已付款，或尚未蓋章）。")
        else:
            st.dataframe(notify_view, use_container_width=True, hide_index=True)
    else:
        st.caption("已過天數給你決定要不要去 SiteGiant 取消訂單；本頁不取消。")


def _render_preorder_vendor_import(campaign_df):
    st.markdown("### 1. 從廠商單帶入")
    st.caption(
        "勾選要跟的款再寫入。只帶入廠商檔名、貨號、品名、條碼；不會填庫存 SKU。"
        " 沒條碼只要勾選也會匯入。已有列的 SKU、自購、上限會保留。"
        " 已填實際關閉的列不會被這次匯入覆寫。"
    )
    uploaded = st.file_uploader(
        "上傳廠商原始訂購單（xlsx／xls）",
        type=["xlsx", "xls"],
        key="preorder_vendor_po_upload",
        help="支援 .xlsx 與 .xls。測用如 2026-License(SJ0910).xlsx。",
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
    st.caption(
        f"檔名：`{uploaded.name}`　可勾選 {len(preview)} 列。"
        " 預設不全選，避免整份廠商目錄誤匯入（麗嬰型可能上百列）；請只勾要跟的款，或用下面全選/取消全選輔助。"
    )
    sel_col1, sel_col2 = st.columns(2)
    with sel_col1:
        if st.button("全選", key="preorder_vendor_select_all", use_container_width=True):
            st.session_state.pop("preorder_vendor_preview_editor", None)
            st.session_state["preorder_vendor_preview_override"] = True
            st.rerun()
    with sel_col2:
        if st.button("取消全選", key="preorder_vendor_select_none", use_container_width=True):
            st.session_state.pop("preorder_vendor_preview_editor", None)
            st.session_state["preorder_vendor_preview_override"] = False
            st.rerun()
    override = st.session_state.pop("preorder_vendor_preview_override", None)
    if override is not None:
        preview = preview.copy()
        preview["勾選"] = override
    edited_preview = st.data_editor(
        preview,
        hide_index=True,
        use_container_width=True,
        disabled=["貨號", "品名", "條碼", "sheet", "row"],
        column_config={
            "勾選": st.column_config.CheckboxColumn("勾選", default=False),
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
        st.session_state.pop("preorder_lock_preview_self_buy", None)
        st.session_state.pop("preorder_lock_result", None)
        st.session_state.pop("preorder_lock_selected", None)
        _clear_preorder_lock_select_editor()
        st.session_state["preorder_vendor_import_msg"] = (
            f"已寫入活動表：新增 {result['added']} 列、更新 {result['updated']} 列。"
        )
        skipped_closed = result.get("skipped_closed") or 0
        if skipped_closed:
            st.session_state["preorder_vendor_import_msg"] += f" {skipped_closed} 列已實際關閉，未覆寫。"
        for note in result.get("reports") or []:
            st.session_state["preorder_vendor_import_msg"] += " " + note
        st.rerun()
    return campaign_df


def _selected_lock_indices(edited_df):
    if edited_df is None or getattr(edited_df, "empty", True):
        return []
    selected = []
    for rec in edited_df.to_dict(orient="records"):
        if rec.get("勾選"):
            selected.append(int(rec.get("index")))
    return selected


def _lock_self_buy_by_index(edited_df):
    if edited_df is None or getattr(edited_df, "empty", True) or "index" not in edited_df.columns:
        return {}
    out = {}
    for rec in edited_df.to_dict(orient="records"):
        try:
            idx = int(rec.get("index"))
        except (TypeError, ValueError):
            continue
        out[idx] = _preorder_self_buy(rec.get("自購"))
    return out


def _lock_self_buy_dirty(edited_df, campaign_df):
    edited_map = _lock_self_buy_by_index(edited_df)
    if not edited_map:
        return False
    campaign = ensure_preorder_campaign_df(campaign_df)
    for idx, buy in edited_map.items():
        if idx not in campaign.index:
            return True
        camp_buy = _preorder_self_buy(campaign.at[idx, "自購"] if "自購" in campaign.columns else 0)
        if camp_buy != buy:
            return True
    return False


def _apply_lock_editor_self_buy(campaign_df, edited_df):
    campaign = ensure_preorder_campaign_df(campaign_df).copy()
    if "自購" not in campaign.columns:
        campaign["自購"] = 0
    for idx, buy in _lock_self_buy_by_index(edited_df).items():
        if idx not in campaign.index:
            continue
        campaign.at[idx, "自購"] = buy
    return campaign


def _lock_select_index_set(df):
    if df is None or getattr(df, "empty", True) or "index" not in getattr(df, "columns", []):
        return set()
    return {int(i) for i in df["index"].tolist()}


def _render_preorder_close(campaign_df, orders_df):
    campaign_df = st.session_state.get("preorder_campaign_df", campaign_df)
    st.markdown("### 4. 結單下載")
    st.caption(
        "勾選要結單的列再預覽。鎖定數量 = 自購 + min(客戶量, 上限)。上限空＝不封頂；自購空＝0。"
        " 未填 SKU、已實際關閉的列不會出現在勾選表。"
        " 自購可在這張表改；按「預覽鎖定數量」才寫入活動表（尚未覆寫雲端）。"
        " 勾選表上的鎖定數量以預覽為準。"
        " SiteGiant 採購單只給本機下載，不會覆寫雲端空殼。"
        " 同場再確認鎖定，會取代本次尚未存檔的廠商單歷史。"
    )
    dups = duplicate_preorder_skus(campaign_df)
    if dups:
        st.warning("同一 SKU 出現在多列，客戶量會重複加總：" + "、".join(f"`{s}`" for s in dups))

    fill_close = st.checkbox("鎖定時把「實際關閉」填成今天", value=True, key="preorder_fill_close_date")
    close_date = taipei_now().strftime("%Y-%m-%d") if fill_close else None
    lock_result = st.session_state.get("preorder_lock_result")
    preview = st.session_state.get("preorder_lock_preview")
    candidates = preorder_lock_selection_df(campaign_df, orders_df)
    edited_candidates = None
    selected_indices = []
    existing_base = st.session_state.get("preorder_lock_select_base")
    if existing_base is not None and _lock_select_index_set(existing_base) != _lock_select_index_set(
        candidates
    ):
        _clear_preorder_lock_select_editor()

    if candidates.empty:
        st.info("沒有可結單的列（需要已填 SKU、且尚未實際關閉）。")
    else:
        st.caption(f"可結單 {len(candidates)} 列。取消勾選的列這次不鎖定。自購改完請按預覽才寫入活動表。")
        sel_col1, sel_col2 = st.columns(2)
        with sel_col1:
            if st.button("全選", key="preorder_lock_select_all", use_container_width=True):
                _clear_preorder_lock_select_editor()
                st.session_state["preorder_lock_select_override"] = True
                st.rerun()
        with sel_col2:
            if st.button("取消全選", key="preorder_lock_select_none", use_container_width=True):
                _clear_preorder_lock_select_editor()
                st.session_state["preorder_lock_select_override"] = False
                st.rerun()
        lock_override = st.session_state.pop("preorder_lock_select_override", None)
        if lock_override is not None:
            candidates = candidates.copy()
            candidates["勾選"] = lock_override
            st.session_state["preorder_lock_select_base"] = candidates.copy()
        lock_base = _editor_base("preorder_lock_select_base", candidates)
        edited_candidates = st.data_editor(
            lock_base,
            hide_index=True,
            use_container_width=True,
            disabled=["index", "品名", "SKU", "條碼", "貨號", "客戶量", "上限", "鎖定數量"],
            column_config={
                "勾選": st.column_config.CheckboxColumn("結單", default=True),
                "index": st.column_config.NumberColumn("列", disabled=True),
                "自購": st.column_config.NumberColumn(
                    "自購", min_value=0, step=1, help="按預覽鎖定數量後才寫入活動表"
                ),
            },
            column_order=["勾選", "品名", "SKU", "自購", "客戶量", "上限", "鎖定數量", "條碼", "貨號"],
            key="preorder_lock_select_editor",
        )
        selected_indices = _selected_lock_indices(edited_candidates)
        if not lock_result:
            st.caption(f"目前勾選 {len(selected_indices)} 列。")
        if _lock_self_buy_dirty(edited_candidates, campaign_df):
            st.warning("自購已改、尚未套用到活動表。請按預覽鎖定數量；不要先按第 2 段覆寫雲端。")

    if st.button("預覽鎖定數量", use_container_width=True, key="preorder_lock_preview_btn"):
        if not selected_indices:
            st.warning("沒有勾選任何可結單列。")
        else:
            campaign_df = _apply_lock_editor_self_buy(campaign_df, edited_candidates)
            st.session_state["preorder_campaign_df"] = campaign_df
            _clear_preorder_campaign_editor()
            st.session_state["preorder_lock_selected"] = list(selected_indices)
            st.session_state["preorder_lock_preview_self_buy"] = _lock_self_buy_by_index(
                edited_candidates
            )
            st.session_state["preorder_lock_preview"] = lock_preorder_campaign(
                campaign_df, orders_df, close_date=close_date, selected_indices=selected_indices
            )
            st.session_state.pop("preorder_lock_result", None)
            st.session_state.pop("preorder_vendor_fill", None)
            st.session_state.pop("preorder_restock_fill", None)
            st.rerun()

    if preview and not lock_result:
        preview_selected = st.session_state.get("preorder_lock_selected") or []
        st.info(
            f"以下為預覽（{len(preview.get('locked_rows') or [])} 列），尚未寫入廠商單歷史。"
            " 核對後再按確認鎖定。若改了勾選或自購請再按預覽。"
        )
        selection_changed = set(preview_selected) != set(selected_indices)
        self_buy_changed = st.session_state.get(
            "preorder_lock_preview_self_buy"
        ) != _lock_self_buy_by_index(edited_candidates)
        if selection_changed or self_buy_changed:
            st.warning("勾選或自購已改，請再按預覽鎖定數量。")
        _show_preorder_lock_tables(preview)
        if (not selection_changed) and (not self_buy_changed) and st.button(
            "確認鎖定", type="primary", use_container_width=True, key="preorder_lock_confirm_btn"
        ):
            campaign_df = _apply_lock_editor_self_buy(campaign_df, edited_candidates)
            st.session_state["preorder_campaign_df"] = campaign_df
            result = lock_preorder_campaign(
                campaign_df, orders_df, close_date=close_date, selected_indices=preview_selected
            )
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
                if filled.get("ok"):
                    filled["lock_sync"] = sync_preorder_vendor_lock_record(
                        filled.get("bytes"), vendor_name
                    )
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
                sync = filled.get("lock_sync") or {}
                if sync.get("ok"):
                    mode = sync.get("mode") or ""
                    if mode == "archive":
                        st.caption(
                            f"☁️ 已併入雲端彙整「{sync.get('filename')}」"
                            f"（本次：`{sync.get('member') or ''}`；個人雲端無法逐檔新建時用此方式）"
                        )
                    else:
                        st.caption(f"☁️ 已同步更新雲端「結單紀錄」：`{sync.get('filename')}`")
                else:
                    st.warning(
                        f"⚠️ 結單紀錄同步雲端失敗：{sync.get('reason') or '未知錯誤'}"
                        "（不影響本機下載；改了勾選再按一次「預覽鎖定數量」／「確認鎖定」會重試）"
                    )
                campaign_miss = filled.get("campaign_unmatched") or []
                vendor_skip = filled.get("vendor_unmatched") or []
                if campaign_miss:
                    st.warning(
                        f"有 {len(campaign_miss)} 筆已鎖定列對不到廠商檔（沒有填假數量），請核對條碼／貨號。"
                    )
                    st.dataframe(pd.DataFrame(campaign_miss), hide_index=True, use_container_width=True)
                if vendor_skip:
                    with st.expander(
                        f"廠商檔其他品項（{len(vendor_skip)} 列，未鎖定／未開放預購，可忽略）",
                        expanded=False,
                    ):
                        st.caption("這些列在開單時未納入客戶預購或未勾選鎖定，回填時略過屬正常。")
                        st.dataframe(pd.DataFrame(vendor_skip), hide_index=True, use_container_width=True)
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
                    _clear_inward_grid_editor()
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
            _editor_base("inward_grid_base", st.session_state["inward_input_df"]),
            num_rows="dynamic",
            use_container_width=True,
            key="inward_grid",
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
                    _clear_inward_grid_editor()
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
        st.caption(
            "重抓雲端活動表，並清空本頁所有未存檔的編輯（廠商檔快取、結單預覽、All Orders 快取等）。"
            " 只想重抓訂單看板請用下方「3. 對客戶訂單」裡的『重新載入雲端 All Orders』。"
        )
        if reload or "preorder_loaded" not in st.session_state:
            if reload:
                get_cached_gdrive_file_bytes.clear()
                _clear_preorder_campaign_editor()
                st.session_state.pop("preorder_orders_loaded", None)
                st.session_state.pop("preorder_lock_result", None)
                st.session_state.pop("preorder_lock_preview", None)
                st.session_state.pop("preorder_lock_preview_self_buy", None)
                st.session_state.pop("preorder_lock_selected", None)
                _clear_preorder_lock_select_editor()
                st.session_state.pop("preorder_vendor_history_before_lock", None)
                _clear_preorder_sku_overview_state()
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

        st.markdown("### 🗺️ 生命週期總覽")
        _overview_campaign_df = st.session_state.get("preorder_campaign_df", loaded["campaign"])
        if "preorder_orders_loaded" not in st.session_state:
            with st.spinner("⏳ 正在載入雲端最新 All Orders…"):
                st.session_state["preorder_orders_loaded"] = load_preorder_orders()
        _cached_orders_loaded = st.session_state.get("preorder_orders_loaded") or {}
        _cached_orders_df = (
            _cached_orders_loaded.get("orders") if _cached_orders_loaded.get("ok") else None
        )
        if _cached_orders_df is not None:
            _overview_board = build_preorder_board(_overview_campaign_df, _cached_orders_df)
            _render_preorder_lifecycle_overview(
                _overview_campaign_df,
                st.session_state.get("preorder_sku_status_df"),
                unpaid_skus_from_board(_overview_board),
                sku_accepted_qty_map(_overview_board),
            )
        else:
            _render_preorder_lifecycle_overview(
                _overview_campaign_df,
                st.session_state.get("preorder_sku_status_df"),
            )
                   
        edited = st.session_state.get("preorder_campaign_df", loaded["campaign"])
        edited = _render_preorder_vendor_import(edited)
        st.session_state["preorder_campaign_df"] = edited

        st.markdown("### 2. 補賣場資料並存檔")
        _render_preorder_campaign_editor()
        edited = st.session_state.get("preorder_campaign_df", loaded["campaign"])

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
