import streamlit as st
import pandas as pd

from utils import collect_tracked_file_status, format_gdrive_time, taipei_now


@st.cache_data(ttl=300, show_spinner="☁️ 正在掃描雲端各資料夾最新檔…")
def _cached_status():
    return collect_tracked_file_status()


def render(sub_page, cfg):
    st.title(sub_page)
    st.info("目前導覽路徑： 📡 雲端資料狀態 ➔ 📋 各表最新修改時間")
    st.write("---")

    st.caption(
        f"監看 Drive 各來源最新檔與台北時間。未接入流程的蝦皮資料夾只顯示狀態、不會寫入主表。"
        f" 現在（台北）：`{taipei_now().strftime('%Y-%m-%d %H:%M:%S')}`"
    )
    if cfg.get("TIME_MASTER"):
        st.caption(f"麗嬰總表快取時間：`{format_gdrive_time(cfg.get('TIME_MASTER'))}`")

    col_btn, _ = st.columns([1, 3])
    with col_btn:
        if st.button("🔄 重新整理狀態", type="primary", use_container_width=True):
            _cached_status.clear()
            st.rerun()

    try:
        rows = _cached_status()
    except Exception:
        st.error("掃描雲端狀態失敗（憑證缺失／無效或資料夾無權限）。")
        return

    df = pd.DataFrame(rows)
    st.dataframe(df, use_container_width=True, hide_index=True)

    consumed = sum(1 for r in rows if r.get("用途") == "已接入流程")
    watch = len(rows) - consumed
    st.caption(f"已接入 `{consumed}` 個來源，僅監看 `{watch}` 個。")

    with st.expander("📧 每日郵件摘要（尚未啟用）"):
        st.markdown(
            "定時同步已會把同一份狀態寫進 GitHub Actions 工作摘要。"
            " 若要寄信，請在 repo Secrets 加上 `STATUS_MAIL_TO`、`SMTP_HOST`、`SMTP_PORT`、"
            "`SMTP_USER`、`SMTP_PASSWORD`、`SMTP_FROM`；沒有這些機密時不會寄信，也不會寫入程式碼。"
        )
