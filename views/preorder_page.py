import sys
import os

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import streamlit as st
import pandas as pd

from utils import (
    get_cached_gdrive_file_bytes,
    clean_barcode,
    format_gdrive_time,
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
    cutoff_to_datetime,
    format_preorder_cutoff_datetime,
    is_preorder_row_closed,
)

from sku_submit_client import (
    category_code_options,
    fetch_sku_form_meta,
    sku_submit_configured,
    suggest_from_name,
    submit_pending_batch,
)


PREORDER_SUBPAGES = (
    "總覽",
    "建檔／補資料",
    "對訂單",
    "結單",
    "到貨催款",
)

PREORDER_CAMPAIGN_UI_COLUMN_ORDER = (
    "月份",
    "結單日",
    "SKU",
    "SKU審核",
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
    st.session_state.pop("preorder_campaign_editor_open_v2", None)
    st.session_state.pop("preorder_campaign_editor_closed_v2", None)
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

def _clear_preorder_sku_overview_state():
    """重新載入雲端資料後，清掉「對訂單」彙總表/搜尋框/明細選取，
    以及「到貨催款」標到貨編輯表的殘留 widget 狀態。"""
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
        _apply_preorder_save_success(result)
        st.rerun()
    st.error(f"❌ 寫回失敗：{result.get('reason') or '未知錯誤'}")


def _apply_preorder_save_success(result):
    """存檔成功後更新 session，並記日曆 warning（若有）。"""
    if result.get("campaign") is not None:
        st.session_state["preorder_campaign_df"] = result["campaign"]
    warning = result.get("calendar_warning")
    if warning:
        st.session_state["preorder_calendar_warning"] = warning
    else:
        st.session_state.pop("preorder_calendar_warning", None)
    st.session_state["preorder_save_ok"] = True
    cal = result.get("calendar") or {}
    bits = []
    if cal.get("created"):
        bits.append(f"新建 {cal['created']}")
    if cal.get("updated"):
        bits.append(f"更新 {cal['updated']}")
    if cal.get("deleted"):
        bits.append(f"刪除 {cal['deleted']}")
    if bits and not warning:
        st.session_state["preorder_calendar_ok"] = "日曆：" + "、".join(bits)
    elif not warning:
        st.session_state.pop("preorder_calendar_ok", None)
    get_cached_gdrive_file_bytes.clear()
    st.session_state.pop("preorder_loaded", None)
    # 清編輯器 base，避免舊「日曆事件ID」空白把剛同步的 ID 蓋掉造成重複建事件
    _clear_preorder_campaign_editor()


def _persist_preorder_tracker(campaign_df):
    cutoff_errors = validate_preorder_campaign_cutoffs(campaign_df)
    if cutoff_errors:
        shown = cutoff_errors[:12]
        reason = (
            "結單日驗證未通過（有 SKU 的列必填，請用結單日欄的日期時間挑選器）：\n"
            + "\n".join(shown)
        )
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

# 共用短句：錯誤／空狀態「問題 → 下一步」
MSG_LOAD_ORDERS = "請先到「對訂單」載入 All Orders。"
MSG_ORDERS_LOAD_FAIL = "無法載入 All Orders。可改上方本機上傳，或檢查資料夾權限與檔名。"
MSG_TRACKER_LOAD_FAIL = "無法載入預購追蹤。請確認該檔有編輯權，且為 .xlsx。"


def _render_preorder_lifecycle_overview(
    campaign_df, sku_status_df, unpaid_skus=None, accepted_qty_by_sku=None
):
    """頁面最上方的生命週期總覽：依「實際關閉」＋「SKU狀態」把每個活動列分到 4 個階段。"""
    staged = compute_preorder_campaign_stages(campaign_df, sku_status_df, unpaid_skus)
    if staged.empty:
        st.caption("活動表沒有列。")
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
        captions.append(f"催款／完成需訂單資料。{MSG_LOAD_ORDERS}")
    if accepted_qty_by_sku is None:
        captions.append(f"「已結單待到貨」尚未排除零訂單 SKU。{MSG_LOAD_ORDERS}")
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
        "結單日": st.column_config.DatetimeColumn(
            "結單日",
            help="有 SKU 必填；用挑選器選日期時間（寫回為 YYYY-MM-DD HH:mm）",
            format="YYYY-MM-DD HH:mm",
            step=60,
            disabled=d("結單日"),
        ),
        "SKU": st.column_config.TextColumn(
            "SKU",
            help="自定義編碼／庫存 SKU；四款請填四個不同 SKU。匯入廠商單不會填這欄；請用下方「送交待審核」產生。",
            disabled=d("SKU"),
        ),
        "SKU審核": st.column_config.TextColumn(
            "SKU審核",
            help="送審後為「待審核」；老闆核准進正式表後可改為「已入正式表」或清空。",
            disabled=d("SKU審核"),
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


def _campaign_df_for_cutoff_editor(df):
    """data_editor 前：結單日轉成 datetime／NaT，供 DatetimeColumn。"""
    out = df.copy() if df is not None else df
    if out is None or getattr(out, "empty", True) or "結單日" not in getattr(out, "columns", []):
        return out
    out = out.copy()
    out["結單日"] = out["結單日"].map(cutoff_to_datetime)
    return out


def _campaign_df_from_cutoff_editor(df):
    """data_editor 後：結單日格式化回 YYYY-MM-DD HH:mm 字串再 ensure。"""
    out = df.copy() if df is not None else df
    if out is not None and not getattr(out, "empty", True) and "結單日" in getattr(out, "columns", []):
        out = out.copy()
        out["結單日"] = out["結單日"].map(format_preorder_cutoff_datetime)
    return ensure_preorder_campaign_df(out)


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

    st.caption("進行中可編輯；填實際關閉會移到下方並鎖定自購／上限／SKU。細節見各欄 help。")

    open_for_editor = _campaign_df_for_cutoff_editor(open_rows)
    open_base = _editor_base("preorder_editor_open_base", open_for_editor, replace_empty=True)
    open_base = _campaign_df_for_cutoff_editor(open_base)
    st.session_state["preorder_editor_open_base"] = open_base
    edited_open = st.data_editor(
        open_base,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_order=list(PREORDER_CAMPAIGN_UI_COLUMN_ORDER),
        column_config=_preorder_campaign_column_config(),
        key="preorder_campaign_editor_open_v2",
    )
    edited_open = _campaign_df_from_cutoff_editor(edited_open)

    with st.expander(
        f"🔒 已關閉（{len(closed_rows)} 列，自購／上限／SKU／月份／結單日已鎖定）",
        expanded=False,
    ):
        if closed_rows.empty:
            st.caption("目前沒有已關閉的活動列。")
            edited_closed = closed_rows
        else:
            st.caption("已關閉列鎖定欄位不可改；清空實際關閉可移回進行中。")
            closed_for_editor = _campaign_df_for_cutoff_editor(closed_rows)
            closed_base = _editor_base(
                "preorder_editor_closed_base", closed_for_editor, replace_empty=True
            )
            closed_base = _campaign_df_for_cutoff_editor(closed_base)
            st.session_state["preorder_editor_closed_base"] = closed_base
            edited_closed = st.data_editor(
                closed_base,
                num_rows="fixed",
                use_container_width=True,
                hide_index=True,
                column_order=list(PREORDER_CAMPAIGN_UI_COLUMN_ORDER),
                column_config=_preorder_campaign_column_config(PREORDER_CLOSED_LOCKED_COLUMNS),
                key="preorder_campaign_editor_closed_v2",
            )
            edited_closed = _campaign_df_from_cutoff_editor(edited_closed)

    open_became_closed = int(_preorder_closed_flags(edited_open).sum())
    closed_became_open = int((~_preorder_closed_flags(edited_closed)).sum()) if len(edited_closed) else 0

    merged = ensure_preorder_campaign_df(pd.concat([edited_open, edited_closed], ignore_index=True))
    st.session_state["preorder_campaign_df"] = merged

    if open_became_closed or closed_became_open:
        st.session_state.pop("preorder_campaign_editor_open_v2", None)
        st.session_state.pop("preorder_campaign_editor_closed_v2", None)
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
    st.markdown("### 對客戶訂單")
    st.caption("摘要看板與 SKU 狀態（唯讀）。")

    reload_orders = st.button(
        "🔄 重新載入雲端 All Orders",
        use_container_width=True,
        key="preorder_orders_reload",
        help="只重抓 All Orders 與訂單看板；不影響建檔未存檔編輯、廠商檔快取或結單預覽。",
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
            "上傳 Orders zip／xlsx",
            type=["xlsx", "xls", "zip"],
            key="preorder_orders_local",
            help="可上傳 Orders_DD-MM-YYYY-*.zip 或已解開的 xlsx。僅本機預覽，不會寫入 Drive。",
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
        st.info(MSG_ORDERS_LOAD_FAIL)
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
        st.warning("活動表沒有填 SKU，兩欄對不到預購列。請到「建檔／補資料」把 SKU 填成預購列的庫存SKU。")
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
    """「對訂單」：可搜尋的彙總表＋選取才渲染明細。狀態欄預設唯讀；
    要標到貨請到「到貨催款」。"""
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
                "到貨 SKU 不在活動表（不自動寫入；要改狀態請到「到貨催款」）："
                + "、".join(f"`{s}`" for s in missing)
            )
        if over_active_skus:
            st.warning("⚠️ 已接單已達或超過上限：" + "、".join(f"`{sku}`" for sku in over_active_skus))

        overview = build_preorder_sku_overview(board, sku_status, stage_by_sku)
        if overview.empty:
            st.info("目前沒有 SKU。")
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
                st.caption("改狀態後按下方覆寫雲端才存檔。")
                _save_preorder_tracker_button(campaign_df, "preorder_save_sku_overview", primary=False)
            else:
                edited_overview = filtered
                st.dataframe(filtered, use_container_width=True, hide_index=True)
                st.caption("狀態唯讀；標到貨請到「到貨催款」。")

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
                st.info("尚無訂單明細（All Orders 未載入）。")
            else:
                st.caption("選取 SKU 後顯示明細。")

    with tab_raw:
        st.caption("All Orders 預購列快照（稽核用，口徑可能與看板不同）。")
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
    """「到貨催款」開頭：唯一可改「預購／到貨」的地方。"""
    st.markdown("**標到貨**")
    st.caption("改「預購／到貨」後按覆寫雲端；下方催款／可打單會依畫面重算。")
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
        st.info("目前沒有可標到貨的 SKU。")
        return sku_status

    # 只列尚未標到貨（狀態＝預購）且已接單（客戶量）＞0；已到貨／零接單不出現在這張編輯表。
    pending_arrival = overview.loc[
        overview["狀態"].map(_preorder_text) == PREORDER_STATUS_PREORDER
    ].reset_index(drop=True)
    if not pending_arrival.empty and "已接單" in pending_arrival.columns:
        qty = pd.to_numeric(pending_arrival["已接單"], errors="coerce").fillna(0)
        pending_arrival = pending_arrival.loc[qty > 0].reset_index(drop=True)
    awaiting_n = (
        int((pending_arrival["階段"] == PREORDER_STAGE_AWAITING_ARRIVAL).sum())
        if "階段" in pending_arrival.columns and not pending_arrival.empty
        else 0
    )
    st.caption(f"待標到貨 {len(pending_arrival)} 列（已結單待到貨 {awaiting_n}）。")
    if pending_arrival.empty:
        if orders_df is None:
            st.info(MSG_LOAD_ORDERS)
        else:
            st.info("沒有待標到貨的 SKU。")
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
        _apply_preorder_save_success(save_result)
        st.session_state.pop("preorder_save_ok", None)  # 催款用自有 msg，不顯示通用存檔成功
    else:
        msg += " 畫面已有日期，但雲端尚未寫回：" + (save_result.get("reason") or "未知錯誤")
    if save_result.get("calendar_warning"):
        st.session_state["preorder_calendar_warning"] = save_result["calendar_warning"]
    st.session_state["preorder_notify_msg"] = msg
    st.rerun()


def _render_preorder_arrival_copy(campaign_df, orders_df, sku_status_df):
    st.markdown("### 到貨催款")
    st.caption("先標到貨再催款／打單；打單與取消仍在 SiteGiant。規則見上方摺疊。")
    sku_status_df = _render_preorder_arrival_status_editor(
        campaign_df, orders_df, sku_status_df
    )
    st.write("---")
    if orders_df is None:
        st.info(MSG_LOAD_ORDERS)
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
            st.caption("已有首次通知日的列預設不勾；全選可刻意全勾（不覆寫既有日期）。")
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
        st.caption("只列仍未付款；已過天數供判斷是否去 SiteGiant 取消（本頁不取消）。")
        if notify_view is None or getattr(notify_view, "empty", True):
            st.info("目前沒有仍未付款的通知紀錄（可能都已付款，或尚未蓋章）。")
        else:
            st.dataframe(notify_view, use_container_width=True, hide_index=True)
    else:
        st.caption("尚無通知紀錄。")


def _render_preorder_vendor_import(campaign_df):
    st.markdown("### 從廠商單帶入")
    st.caption("勾選要跟的款再寫入；不填庫存 SKU，已關閉列不會覆寫。")
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
    st.caption(f"`{uploaded.name}`：可勾選 {len(preview)} 列（預設不全選）。")
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


def _render_preorder_sku_pending_review(campaign_df):
    """空 SKU 列：審核分類後送交待審核，回填碼並標 SKU審核=待審核。"""
    st.markdown("### 送交待審核（空 SKU）")
    st.caption(
        "確認第二／三／四分類與產品關鍵字後按「送交待審核」。狀態固定預購；"
        "成功後回填 SKU 並標「待審核」。老闆核准正式表後再改標籤。"
    )
    campaign = ensure_preorder_campaign_df(campaign_df)
    if "SKU審核" not in campaign.columns:
        campaign["SKU審核"] = ""

    if not sku_submit_configured():
        st.info(
            "尚未設定 `SKU_SUBMIT_WEBAPP_URL`（`.streamlit/secrets.toml`）。"
            "可先手填 SKU，或設定 Web App `/exec` 網址後再批次送審。"
            "若 Script Properties 有 `BATCH_API_TOKEN`，請一併設 `SKU_SUBMIT_BATCH_TOKEN`。"
        )
        return campaign

    rows = []
    for idx, camp in campaign.iterrows():
        if is_preorder_row_closed(camp):
            continue
        if _preorder_text(camp.get("SKU")):
            continue
        pname = _preorder_text(camp.get("品名"))
        if not pname:
            continue
        rows.append(
            {
                "送審": True,
                "index": int(idx),
                "品名": pname,
                "條碼": clean_barcode(camp.get("條碼")),
                "貨號": _preorder_text(camp.get("貨號")),
                "第二分類": "",
                "第三分類": "",
                "第四分類": "NEW",
                "產品關鍵字": "",
            }
        )

    if not rows:
        st.caption("沒有待補 SKU 的進行中列。匯入廠商單後空 SKU 列會出現在這裡。")
        return campaign

    meta = fetch_sku_form_meta()
    if not meta.get("ok"):
        st.warning(meta.get("message") or "無法載入分類對應表；仍可手填分類代碼後送審。")
    l2_opts = category_code_options(meta, "l2")
    l3_opts = category_code_options(meta, "l3")
    l4_opts = category_code_options(meta, "l4") or ["NEW", "SEC"]

    base = pd.DataFrame(rows)
    override = st.session_state.pop("preorder_sku_review_df", None)
    if isinstance(override, pd.DataFrame) and not override.empty:
        # keep用户 edits when re-running suggest
        by_idx = {int(r["index"]): r for r in override.to_dict(orient="records")}
        merged = []
        for r in rows:
            prev = by_idx.get(int(r["index"]))
            if prev:
                r = {
                    **r,
                    "送審": bool(prev.get("送審", True)),
                    "第二分類": _preorder_text(prev.get("第二分類")) or r["第二分類"],
                    "第三分類": _preorder_text(prev.get("第三分類")) or r["第三分類"],
                    "第四分類": _preorder_text(prev.get("第四分類")) or r["第四分類"],
                    "品名": _preorder_text(prev.get("品名")) or r["品名"],
                    "產品關鍵字": _preorder_text(prev.get("產品關鍵字")),
                    "條碼": clean_barcode(prev.get("條碼"))
                    if prev.get("條碼") is not None
                    else r["條碼"],
                }
            merged.append(r)
        base = pd.DataFrame(merged)

    col_cfg = {
        "送審": st.column_config.CheckboxColumn("送審", default=True),
        "index": st.column_config.NumberColumn("index", disabled=True),
        "貨號": st.column_config.TextColumn("貨號", disabled=True),
        "條碼": st.column_config.TextColumn("條碼"),
        "品名": st.column_config.TextColumn("品名"),
        "產品關鍵字": st.column_config.TextColumn(
            "產品關鍵字", help="可空；「依品名建議分類」可預填，送審前可改"
        ),
        "第四分類": st.column_config.SelectboxColumn("第四分類", options=l4_opts, required=True),
    }
    if l2_opts:
        col_cfg["第二分類"] = st.column_config.SelectboxColumn(
            "第二分類", options=[""] + l2_opts, required=True
        )
    if l3_opts:
        col_cfg["第三分類"] = st.column_config.SelectboxColumn(
            "第三分類", options=[""] + l3_opts, required=True
        )

    edited = st.data_editor(
        base,
        hide_index=True,
        use_container_width=True,
        disabled=["index", "貨號"],
        column_config=col_cfg,
        column_order=[
            "送審",
            "品名",
            "條碼",
            "貨號",
            "第二分類",
            "第三分類",
            "第四分類",
            "產品關鍵字",
            "index",
        ],
        key="preorder_sku_pending_editor",
    )

    b1, b2 = st.columns(2)
    with b1:
        if st.button("依品名建議分類", use_container_width=True, key="preorder_sku_suggest_btn"):
            suggested_rows = []
            for rec in edited.to_dict(orient="records"):
                if not rec.get("送審"):
                    suggested_rows.append(rec)
                    continue
                name = _preorder_text(rec.get("品名"))
                sug = suggest_from_name(name) if name else {}
                if sug.get("ok"):
                    rec = dict(rec)
                    if sug.get("第二分類"):
                        rec["第二分類"] = sug["第二分類"]
                    if sug.get("第三分類"):
                        rec["第三分類"] = sug["第三分類"]
                    if sug.get("第四分類"):
                        rec["第四分類"] = sug["第四分類"]
                    if sug.get("產品關鍵字") and not _preorder_text(rec.get("產品關鍵字")):
                        rec["產品關鍵字"] = sug["產品關鍵字"]
                suggested_rows.append(rec)
            st.session_state["preorder_sku_review_df"] = pd.DataFrame(suggested_rows)
            st.session_state.pop("preorder_sku_pending_editor", None)
            st.rerun()
    with b2:
        n_sel = sum(1 for rec in edited.to_dict(orient="records") if rec.get("送審"))
        confirm = st.button(
            f"送交待審核（{n_sel} 筆）",
            type="primary",
            use_container_width=True,
            key="preorder_sku_submit_btn",
            disabled=n_sel < 1,
        )

    if not confirm:
        return campaign

    payload = []
    indices = []
    for rec in edited.to_dict(orient="records"):
        if not rec.get("送審"):
            continue
        l2 = _preorder_text(rec.get("第二分類"))
        l3 = _preorder_text(rec.get("第三分類"))
        l4 = _preorder_text(rec.get("第四分類")).upper() or "NEW"
        name = _preorder_text(rec.get("品名"))
        if not name or not l2 or not l3 or l4 not in ("NEW", "SEC"):
            st.error(
                f"第 index={rec.get('index')} 列缺品名或分類（L4 須 NEW／SEC），修正後再送。"
            )
            return campaign
        payload.append(
            {
                "狀態": "預購",
                "品名": name,
                "c": clean_barcode(rec.get("條碼")),
                "第二分類": l2,
                "第三分類": l3,
                "第四分類": l4,
                "產品關鍵字": _preorder_text(rec.get("產品關鍵字")),
            }
        )
        indices.append(int(rec["index"]))

    with st.spinner(f"送交 {len(payload)} 筆待審核…"):
        result = submit_pending_batch(payload)

    if not result.get("ok"):
        st.error(result.get("message") or "送審失敗")
        return campaign

    out_rows = result.get("rows") or []
    if len(out_rows) != len(indices):
        st.error("回傳筆數與送審筆數不符，未寫入活動表 SKU。請查待審核表。")
        return campaign

    for i, camp_idx in enumerate(indices):
        sku = _preorder_text(out_rows[i].get("自定義編碼"))
        if camp_idx not in campaign.index:
            continue
        campaign.at[camp_idx, "SKU"] = sku
        campaign.at[camp_idx, "SKU審核"] = "待審核"

    st.session_state["preorder_campaign_df"] = ensure_preorder_campaign_df(campaign)
    _clear_preorder_campaign_editor()
    st.session_state.pop("preorder_sku_pending_editor", None)
    st.session_state.pop("preorder_sku_review_df", None)
    msg = result.get("message") or f"已送交 {len(out_rows)} 筆待審核"
    codes = "、".join(_preorder_text(r.get("自定義編碼")) for r in out_rows[:8])
    if len(out_rows) > 8:
        codes += "…"
    st.session_state["preorder_sku_submit_msg"] = f"{msg}：{codes}"
    st.rerun()
    return campaign


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
    st.markdown("### 結單下載")
    st.caption("勾選後按預覽鎖定數量；自購改完也要先預覽。公式見上方計算規則。")
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
        st.caption(f"可結單 {len(candidates)} 列。")
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
            st.warning("自購已改，請先按預覽鎖定數量再覆寫雲端。")

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
            f"預覽 {len(preview.get('locked_rows') or [])} 列（尚未寫入歷史）。核對後按確認鎖定。"
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
        st.caption("尚未快取廠商檔。請到「建檔／補資料」再上傳同一份原檔，或下面補傳。")
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



def _session_orders_df():
    """跨子頁共用：優先本機 All Orders 預覽，否則雲端快取。"""
    local = st.session_state.get("preorder_orders_local_cache")
    if local and local.get("ok"):
        return local.get("orders")
    drive = st.session_state.get("preorder_orders_loaded") or {}
    if drive.get("ok"):
        return drive.get("orders")
    return None


def _reload_preorder_tracker():
    get_cached_gdrive_file_bytes.clear()
    _clear_preorder_campaign_editor()
    st.session_state.pop("preorder_orders_loaded", None)
    st.session_state.pop("preorder_orders_local_cache", None)
    st.session_state.pop("preorder_orders_local_key", None)
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


def _ensure_preorder_loaded():
    if "preorder_loaded" not in st.session_state:
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
    return st.session_state.get("preorder_loaded") or {}


def _render_shared_shell():
    """各子頁共用：重載、雲端狀態、計算規則／空殼摺疊。"""
    st.caption("廠商帶入 → 補 SKU → 對訂單 → 結單 → 到貨催款。")
    reload = st.button(
        "重新載入預購追蹤",
        use_container_width=True,
        key="preorder_reload",
        help="重抓雲端活動表並清空本模組未存檔編輯。只想重抓訂單請到「對訂單」按重新載入 All Orders。",
    )
    if reload:
        _reload_preorder_tracker()
        st.rerun()

    loaded = _ensure_preorder_loaded()
    if not loaded.get("ok"):
        st.error(f"❌ 無法載入預購追蹤：{loaded.get('reason') or '未知錯誤'}")
        st.info(MSG_TRACKER_LOAD_FAIL)
        return None

    if st.session_state.pop("preorder_save_ok", False):
        st.success("已覆寫雲端預購追蹤.xlsx。")
    cal_ok = st.session_state.pop("preorder_calendar_ok", None)
    if cal_ok:
        st.success(cal_ok)
    cal_warn = st.session_state.pop("preorder_calendar_warning", None)
    if cal_warn:
        st.warning(cal_warn)
    import_msg = st.session_state.pop("preorder_vendor_import_msg", None)
    if import_msg:
        st.success(import_msg)

    st.caption(
        f"雲端檔：`{loaded.get('name')}`　最後修改：`{format_gdrive_time(loaded.get('modified'))}`"
    )

    with st.expander("計算規則與 SiteGiant 採購單空殼", expanded=False):
        st.markdown(
            "- **鎖定數量**＝自購 + min(客戶量, 上限)；上限空＝不封頂；自購空＝0。\n"
            "- **看板**：商品名稱含「預購」才計件；已付款／未付款只看付款狀態。\n"
            "- **催款**：已標到貨 × 未付款 × 非已取消／已退款。\n"
            "- **可打單**：已標到貨 × 已付款 × 訂單狀態＝待處理。\n"
            "- 寫回只覆寫既有 `預購追蹤.xlsx`（結單紀錄資料夾例外）。"
        )
        restock = probe_sg_restock_template()
        if restock.get("ok"):
            st.download_button(
                label=f"下載採購單空殼（`{restock.get('name') or 'import_restock.xlsx'}`）",
                data=restock.get("bytes") or b"",
                file_name=restock.get("name") or "import_restock.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="preorder_restock_template_dl",
                help="只供本機參考；結單區另產已填檔，不會覆寫雲端空殼。",
            )
        else:
            st.caption(f"無法讀取採購單空殼：{restock.get('reason') or '未知錯誤'}")

    return st.session_state.get("preorder_campaign_df", loaded["campaign"])


def _nav_to(sub_page_label):
    st.session_state["preorder_nav_target"] = sub_page_label
    st.rerun()


def _render_overview(campaign_df):
    st.markdown("### 生命週期總覽")
    if "preorder_orders_loaded" not in st.session_state:
        with st.spinner("⏳ 正在載入雲端最新 All Orders…"):
            st.session_state["preorder_orders_loaded"] = load_preorder_orders()
    cached_orders_df = _session_orders_df()
    if cached_orders_df is not None:
        overview_board = build_preorder_board(campaign_df, cached_orders_df)
        _render_preorder_lifecycle_overview(
            campaign_df,
            st.session_state.get("preorder_sku_status_df"),
            unpaid_skus_from_board(overview_board),
            sku_accepted_qty_map(overview_board),
        )
    else:
        _render_preorder_lifecycle_overview(
            campaign_df,
            st.session_state.get("preorder_sku_status_df"),
        )

    st.write("---")
    st.markdown("### 前往工作區")
    cols = st.columns(4)
    targets = [
        ("建檔／補資料", "廠商帶入＋補 SKU／自購／上限"),
        ("對訂單", "All Orders 看板（狀態唯讀）"),
        ("結單", "鎖定訂量、下載兩檔"),
        ("到貨催款", "標到貨、催款、可打單"),
    ]
    for col, (label, hint) in zip(cols, targets):
        with col:
            if st.button(label, use_container_width=True, key=f"preorder_nav_{label}"):
                _nav_to(label)
            st.caption(hint)


def render(sub_page):
    """預購追蹤第四主模組：依側邊欄子頁渲染。"""
    st.title(f"{sub_page}")
    st.write("---")

    campaign = _render_shared_shell()
    if campaign is None:
        return

    if sub_page == "總覽":
        _render_overview(campaign)
        return

    if sub_page == "建檔／補資料":
        edited = st.session_state.get("preorder_campaign_df", campaign)
        edited = _render_preorder_vendor_import(edited)
        st.session_state["preorder_campaign_df"] = edited
        if st.session_state.get("preorder_vendor_import_msg"):
            st.success(st.session_state.pop("preorder_vendor_import_msg"))
        edited = _render_preorder_sku_pending_review(
            st.session_state.get("preorder_campaign_df", edited)
        )
        st.session_state["preorder_campaign_df"] = edited
        if st.session_state.get("preorder_sku_submit_msg"):
            st.success(st.session_state.pop("preorder_sku_submit_msg"))
        st.markdown("### 補賣場資料並存檔")
        _render_preorder_campaign_editor()
        edited = st.session_state.get("preorder_campaign_df", campaign)
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
        return

    if sub_page == "對訂單":
        _render_preorder_orders_board(
            st.session_state.get("preorder_campaign_df", campaign)
        )
        return

    if sub_page == "結單":
        orders_df = _session_orders_df()
        if orders_df is None and "preorder_orders_loaded" not in st.session_state:
            with st.spinner("⏳ 正在載入雲端最新 All Orders…"):
                st.session_state["preorder_orders_loaded"] = load_preorder_orders()
            orders_df = _session_orders_df()
        if orders_df is None:
            st.info(f"尚無 All Orders。{MSG_LOAD_ORDERS}")
        _render_preorder_close(
            st.session_state.get("preorder_campaign_df", campaign), orders_df
        )
        return

    if sub_page == "到貨催款":
        orders_df = _session_orders_df()
        if orders_df is None and "preorder_orders_loaded" not in st.session_state:
            with st.spinner("⏳ 正在載入雲端最新 All Orders…"):
                st.session_state["preorder_orders_loaded"] = load_preorder_orders()
            orders_df = _session_orders_df()
        _render_preorder_arrival_copy(
            st.session_state.get("preorder_campaign_df", campaign),
            orders_df,
            st.session_state.get("preorder_sku_status_df"),
        )
        return

    st.warning(f"未知子頁：{sub_page}")
