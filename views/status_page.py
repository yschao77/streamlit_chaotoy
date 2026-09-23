import streamlit as st
import pandas as pd

from utils import collect_tracked_file_status, format_gdrive_time, taipei_now


@st.cache_data(ttl=300, show_spinner="☁️ 正在掃描雲端各資料夾最新檔…")
def _cached_status():
    return collect_tracked_file_status()


def render(sub_page, cfg):
    st.title(sub_page)
    st.write("---")

    st.caption(
        f"監看各來源最新檔（台北 `{taipei_now().strftime('%Y-%m-%d %H:%M:%S')}`）。"
        "未接入流程的只顯示狀態。"
    )
    if cfg.get("TIME_MASTER"):
        st.caption(f"麗嬰總表快取：`{format_gdrive_time(cfg.get('TIME_MASTER'))}`")

    col_btn, _ = st.columns([1, 3])
    with col_btn:
        if st.button("重新整理狀態", type="primary", use_container_width=True):
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
    st.caption(f"已接入 `{consumed}` 個，僅監看 `{watch}` 個。")

    with st.expander("每日郵件摘要（尚未啟用）"):
        st.markdown(
            "定時同步會寫入 Actions 工作摘要。"
            "寄信需在 Secrets 設定 `STATUS_MAIL_TO` 與 SMTP 相關項；未設定則不寄信。"
        )
