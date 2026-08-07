import streamlit as st
import pandas as pd
import io
import datetime

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# 從 utils.py 引入所有工具
from utils import (
    HAS_CALAMINE, download_gdrive_file_to_bytes, upload_or_update_gdrive_file,
    get_cached_gdrive_id, format_gdrive_time, list_gdrive_files, get_cached_gdrive_file_bytes,
    load_master_data, load_shopee_data, save_to_master_xlsm, save_to_shopee_master_xlsm,
    calculate_md5, clean_barcode, process_smart_headers, run_cross_matching
)

def render(sub_page, cfg):
    """
    渲染「商品蝦皮麗嬰統整管理」的子頁面
    cfg 包含所有從 app.py 傳來的檔案 ID 與 metadata
    """
    st.title(f"{sub_page}")
    st.info(f"目前導覽路徑： 📦 商品蝦皮麗嬰統整管理 ➔ {sub_page}")
    st.write("---")

    # 內建全域獨立功能：三表 PowerQuery 整合、計算財務指標
    def run_powerquery_and_update_gdrive(df_to_save=None):
        if not (cfg['ID_LOCAL_PROD'] and cfg['ID_SHOPEE_MASTER'] and cfg['ID_MASTER_FILE']):
            st.error("❌ 缺少核心資料主檔案 ID，無法啟動三表整合！")
            return False
            
        try:
            if df_to_save is None:
                engine_kw = {"engine": "calamine"} if HAS_CALAMINE else {}
                df_liying = pd.read_excel(download_gdrive_file_to_bytes(cfg['ID_MASTER_FILE']), sheet_name="麗嬰國際產品總表", **engine_kw)
                df_p = pd.read_excel(download_gdrive_file_to_bytes(cfg['ID_LOCAL_PROD']), sheet_name=0, **engine_kw)
                df_s = pd.read_excel(download_gdrive_file_to_bytes(cfg['ID_SHOPEE_MASTER']), sheet_name="蝦皮商品列表", **engine_kw)
                
                df_liying['條碼'] = df_liying['條碼'].astype(str).str.strip().str.split('.').str[0]
                df_p["自定義編碼"] = df_p["自定義編碼"].astype(str).str.strip().str.split('.').str[0]
                df_s["iSKU"] = df_s["iSKU"].astype(str).str.strip().str.split('.').str[0]
                
                if "商品名稱" in df_p.columns:
                    df_p = df_p.rename(columns={"商品名稱": "內部商品名稱"})
                
                df_merge1 = pd.merge(df_p, df_s[["商品名稱","iSKU", "GTIN", "價格"]], left_on="自定義編碼", right_on="iSKU", how="left")
                df_merge1 = df_merge1.rename(columns={"商品名稱": "蝦皮商品名稱", "GTIN": "蝦皮GTIN", "價格": "蝦皮售價"})
                df_merge1["c"] = df_merge1["c"].astype(str).str.strip().str.split('.').str[0]
                
                df_final = pd.merge(df_merge1, df_liying[["條碼", "零售價", "含稅"]], left_on="c", right_on="條碼", how="left")
                df_final = df_final.rename(columns={"零售價": "麗嬰零售價", "含稅": "麗嬰批發含稅價", "條碼": "麗嬰條碼"})
                df_final["麗嬰商品"] = df_final["麗嬰條碼"].apply(lambda x: None if pd.isna(x) else "v")
                
                for c in ["蝦皮售價", "麗嬰零售價", "麗嬰批發含稅價"]:
                    df_final[c] = pd.to_numeric(df_final[c], errors='coerce')
                    
                df_final["麗嬰零售八折"] = df_final["麗嬰零售價"] * 0.8
                df_final["麗嬰八折比蝦皮貴"] = df_final.apply(lambda r: "v" if (pd.notna(r["麗嬰零售八折"]) and pd.notna(r["蝦皮售價"]) and r["麗嬰零售八折"] > r["蝦皮售價"]) else None, axis=1)
                df_final["麗嬰未稅價"] = df_final["麗嬰批發含稅價"].apply(lambda x: round(x / 1.05, 2) if pd.notna(x) else None)
                df_final["麗嬰稅款"] = df_final.apply(lambda r: round(r["麗嬰批發含稅價"] - r["麗嬰未稅價"], 2) if (pd.notna(r["麗嬰批發含稅價"]) and pd.notna(r["麗嬰未稅價"])) else None, axis=1)
                
                df_to_save = df_final.drop(columns=["iSKU"], errors="ignore")
                st.session_state['pq_result'] = df_to_save
                
            output_stream = io.BytesIO()
            with pd.ExcelWriter(output_stream, engine='openpyxl') as writer:
                df_to_save.to_excel(writer, index=False, sheet_name="商品蝦皮麗嬰價格統整表")
            output_stream.seek(0)
            
            existing_summary_id, _, _ = get_cached_gdrive_id(cfg['ID_PRICE_SUMMARY_FOLDER'], "商品蝦皮麗嬰價格統整表")
            
            if existing_summary_id:
                upload_or_update_gdrive_file(
                    folder_id=cfg['ID_PRICE_SUMMARY_FOLDER'], 
                    file_name="商品蝦皮麗嬰價格統整表.xlsx", 
                    file_bytes=output_stream.getvalue(), 
                    existing_file_id=existing_summary_id
                )
                
                if "gdrive_id_cache" in st.session_state:
                    cache_key = f"{cfg['ID_PRICE_SUMMARY_FOLDER']}_商品蝦皮麗嬰價格統整表"
                    if cache_key in st.session_state["gdrive_id_cache"]:
                        del st.session_state["gdrive_id_cache"][cache_key]
                return True
            else:
                st.error("❌ 雲端不存在『商品蝦皮麗嬰價格統整表』空白主表檔案，無法執行覆寫更新。")
                return False
                
        except Exception as e:
            st.error(f"❌ 自動化整合或回寫雲端時發生異常: {str(e)}")
            return False

    # -------------------------------------------------------------------------
    # 邏輯區塊：根據 sub_page 呈現功能
    # -------------------------------------------------------------------------
    if sub_page == "📊 PowerQuery 三表整合歷史紀錄":
        st.subheader("🔄 三表整合歷史紀錄追蹤")
        hist_pq_files = list_gdrive_files(cfg['ID_PRICE_SUMMARY_FOLDER'])
        if not hist_pq_files:
            st.warning(f"💡 提示：目前雲端資料夾內尚無任何歷史檔案，請至『🧠 PowerQuery 三表整合』執行新建轉換。")
        else:
            file_options = {f['name']: f['id'] for f in hist_pq_files}
            selected_pq_file = st.selectbox("🎯 請選擇欲調閱的歷史整合報告：", list(file_options.keys()))
            
            if selected_pq_file:
                try:
                    target_id = file_options[selected_pq_file]
                    file_bytes = download_gdrive_file_to_bytes(target_id)
                    df_pq_view = pd.read_excel(file_bytes, engine="calamine" if HAS_CALAMINE else None)
                    st.markdown(f"📄 **目前調閱雲端檔案**：`{selected_pq_file}` ｜ 📊 **資料總項數**：`{len(df_pq_view)} 筆`")
                    st.dataframe(df_pq_view, use_container_width=True)
                    st.download_button(label="🔄 重新下載此歷史整合表 (.xlsx)", data=file_bytes.getvalue(), file_name=selected_pq_file, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                except Exception as e:
                    st.error(f"❌ 讀取雲端備份檔案失敗: {str(e)}")

    elif sub_page == "🧠 PowerQuery 執行三表整合":
        st.subheader("🔍 三表數據追蹤")
        
        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("📦 商品列表 (商品iSKU清單)", "已對接" if cfg['ID_LOCAL_PROD'] else "❌ 未偵測到")
            st.caption(f"📅 最後修改時間: \n`{format_gdrive_time(cfg['TIME_PROD'])}`")
        with c2:
            st.metric("🧡 蝦皮資料庫主表", "已對接" if cfg['ID_SHOPEE_MASTER'] else "❌ 未偵測到")
            st.caption(f"📅 最後修改時間: \n`{format_gdrive_time(cfg['TIME_SHOPEE'])}`")
        with c3:
            st.metric("🧸 麗嬰產品總表", "已對接" if cfg['ID_MASTER_FILE'] else "❌ 未偵測到")
            st.caption(f"📅 最後修改時間: \n`{format_gdrive_time(cfg['TIME_MASTER'])}`")

        st.write("---")

        st.subheader("📊 雲端『商品蝦皮麗嬰價格統整表』當前狀態")
        existing_summary_id, existing_summary_time, _ = get_cached_gdrive_id(cfg['ID_PRICE_SUMMARY_FOLDER'], "商品蝦皮麗嬰價格統整表")
        
        if existing_summary_id:
            st.info(f"🟢 雲端已存在統整表檔案 ｜ 📅 最後修改時間：`{format_gdrive_time(existing_summary_time)}`")
        else:
            st.warning("⚠️ 雲端目前尚未建立『商品蝦皮麗嬰價格統整表』，回寫時系統將會自動全新建立。")

        st.write("---")

        if st.button("🛠️ 啟動三表整合與財務指標計算", type="primary", use_container_width=True):
            if not (cfg['ID_LOCAL_PROD'] and cfg['ID_SHOPEE_MASTER'] and cfg['ID_MASTER_FILE']):
                st.error("❌ 無法啟動三表整合！請確認雲端對應資料夾內是否缺少必要的核心資料主檔案。")
            else:
                with st.spinner("正在由雲端載入數據流並進行大數據跨表計算..."):
                    try:
                        engine_kw = {"engine": "calamine"} if HAS_CALAMINE else {}
                        df_liying = pd.read_excel(download_gdrive_file_to_bytes(cfg['ID_MASTER_FILE']), sheet_name="麗嬰國際產品總表", **engine_kw)
                        df_p = pd.read_excel(download_gdrive_file_to_bytes(cfg['ID_LOCAL_PROD']), sheet_name=0, **engine_kw)
                        df_s = pd.read_excel(download_gdrive_file_to_bytes(cfg['ID_SHOPEE_MASTER']), sheet_name="蝦皮商品列表", **engine_kw)
                        
                        df_liying['條碼'] = df_liying['條碼'].astype(str).str.strip().str.split('.').str[0]
                        df_p["自定義編碼"] = df_p["自定義編碼"].astype(str).str.strip().str.split('.').str[0]
                        df_s["iSKU"] = df_s["iSKU"].astype(str).str.strip().str.split('.').str[0]
                        
                        if "商品名稱" in df_p.columns:
                            df_p = df_p.rename(columns={"商品名稱": "內部商品名稱"})
                        
                        df_merge1 = pd.merge(df_p, df_s[["商品名稱","iSKU", "GTIN", "價格"]], left_on="自定義編碼", right_on="iSKU", how="left")
                        df_merge1 = df_merge1.rename(columns={"商品名稱": "蝦皮商品名稱", "GTIN": "蝦皮GTIN", "價格": "蝦皮售價"})
                        df_merge1["c"] = df_merge1["c"].astype(str).str.strip().str.split('.').str[0]
                        
                        df_final = pd.merge(df_merge1, df_liying[["條碼", "零售價", "含稅"]], left_on="c", right_on="條碼", how="left")
                        df_final = df_final.rename(columns={"零售價": "麗嬰零售價", "含稅": "麗嬰批發含稅價", "條碼": "麗嬰條碼"})
                        df_final["麗嬰商品"] = df_final["麗嬰條碼"].apply(lambda x: None if pd.isna(x) else "v")
                        
                        for c in ["蝦皮售價", "麗嬰零售價", "麗嬰批發含稅價"]:
                            df_final[c] = pd.to_numeric(df_final[c], errors='coerce')
                            
                        df_final["麗嬰零售八折"] = df_final["麗嬰零售價"] * 0.8
                        df_final["麗嬰八折比蝦皮貴"] = df_final.apply(lambda r: "v" if (pd.notna(r["麗嬰零售八折"]) and pd.notna(r["蝦皮售價"]) and r["麗嬰零售八折"] > r["蝦皮售價"]) else None, axis=1)
                        df_final["麗嬰未稅價"] = df_final["麗嬰批發含稅價"].apply(lambda x: round(x / 1.05, 2) if pd.notna(x) else None)
                        df_final["麗嬰稅款"] = df_final.apply(lambda r: round(r["麗嬰批發含稅價"] - r["麗嬰未稅價"], 2) if (pd.notna(r["麗嬰批發含稅價"]) and pd.notna(r["麗嬰未稅價"])) else None, axis=1)
                        
                        st.session_state['pq_result'] = df_final.drop(columns=["iSKU"], errors="ignore")
                        st.success("🎉 三表 PowerQuery 交叉聯結與財務指標計算整合完成！")
                    except Exception as e:
                        st.error(f"❌ 錯誤: {str(e)}")

        if 'pq_result' in st.session_state and st.session_state['pq_result'] is not None:
            df_result = st.session_state['pq_result']
            st.subheader("📋 整合聯結情報報表輸出預覽")
            st.markdown(f"📊 **目前整合結果資料總項數**：`{len(df_result)} 筆`")
            st.dataframe(df_result, use_container_width=True)
            
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                towrite_pq = io.BytesIO()
                with pd.ExcelWriter(towrite_pq, engine='openpyxl') as writer:
                    df_result.to_excel(writer, index=False, sheet_name="PowerQuery三表整合")
                st.download_button(
                    label="📥 匯出並下載此三表整合交叉比對表 (.xlsx)", 
                    data=towrite_pq.getvalue(), 
                    file_name=f"三表整合比對結果_{datetime.date.today().strftime('%Y%m%d')}.xlsx", 
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            with col_btn2:
                if st.button("🔄 執行：將整合結果回寫並更新至雲端", type="secondary", use_container_width=True):
                    with st.spinner("💾 正在覆寫更新雲端現有統整表檔案..."):
                        if run_powerquery_and_update_gdrive(df_to_save=df_result):
                            st.success("✅ 雲端統整表已成功同步覆寫更新！")
                            st.info("💡 重新整理頁面後，上方將會顯示最新的修改時間。")

    elif sub_page == "🔍 麗嬰商品總表數據查詢":
        st.subheader("📋 麗嬰採購產品總表資料庫分頁動態檢視")
        
        _, _, _, _, all_sheets, _ = load_master_data(cfg['ID_MASTER_FILE'])
        
        if all_sheets:
            view_sheets = [s for s in all_sheets if s != "麗嬰產品新採購單"]
            selected_sheet = st.selectbox("請選擇數據分頁：", view_sheets)
            search_mode = st.radio("🎯 請選擇查詢模式：", ["多筆條碼價格查詢", "多筆庫存SKU查詢", "模糊關鍵字搜尋"], horizontal=True)
            
            try:
                file_bytes = get_cached_gdrive_file_bytes(cfg['ID_MASTER_FILE'])
                engine_kw = {"engine": "calamine"} if HAS_CALAMINE else {}
                df_view = pd.read_excel(io.BytesIO(file_bytes), sheet_name=selected_sheet, **engine_kw)
                
                if search_mode == "多筆條碼價格查詢":
                    if "條碼" not in df_view.columns:
                        st.warning(f"⚠️ 當前選擇的分頁 【{selected_sheet}】 內部不含「條碼」欄位。")
                    else:
                        df_view['條碼'] = df_view['條碼'].apply(clean_barcode)
                        barcode_paste = st.text_area("📋 請貼上多筆國際條碼 (換行、空格或逗號隔開)：", height=120)
                        
                        if barcode_paste.strip():
                            import re
                            cleaned_barcodes = [clean_barcode(t) for t in re.split(r'[\n,\s]+', barcode_paste) if t.strip()]
                            cleaned_barcodes = [b for b in cleaned_barcodes if b]
                            
                            if cleaned_barcodes:
                                df_result = df_view[df_view['條碼'].isin(cleaned_barcodes)].copy()
                                st.success(f"🔍 查詢完畢！成功比對出 {len(df_result)} 筆商品資料。")
                                
                                important_cols = ["條碼", "名稱", "零售價", "含稅"]
                                display_cols = [c for c in important_cols if c in df_result.columns] + [c for c in df_result.columns if c not in important_cols]
                                st.dataframe(df_result[display_cols], use_container_width=True)
                                
                                towrite_query = io.BytesIO()
                                with pd.ExcelWriter(towrite_query, engine='openpyxl') as writer:
                                    df_result[display_cols].to_excel(writer, index=False, sheet_name="條碼查詢結果")
                                st.download_button("📥 下載本次條碼查詢結果報表 (.xlsx)", data=towrite_query.getvalue(), file_name=f"條碼批次查詢結果_{datetime.date.today().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                
                elif search_mode == "多筆庫存SKU查詢":
                    sku_col_candidate = next((c for c in ["庫存SKU", "自定義編碼", "商品編號", "貨號"] if c in df_view.columns), None)
                    if not sku_col_candidate:
                        st.warning(f"⚠️ 當前選擇的分頁 【{selected_sheet}】 找不到對應的 SKU 欄位。")
                    else:
                        df_view[sku_col_candidate] = df_view[sku_col_candidate].astype(str).str.strip().str.split('.').str[0]
                        sku_paste = st.text_area(f"📋 請貼上多筆【{sku_col_candidate}】（換行、空格或逗號隔開）：", height=120)
                        
                        if sku_paste.strip():
                            import re
                            cleaned_skus = [t.strip() for t in re.split(r'[\n,\s]+', sku_paste) if t.strip()]
                            if cleaned_skus:
                                df_result = df_view[df_view[sku_col_candidate].isin(cleaned_skus)].copy()
                                st.success(f"🔍 查詢完畢！成功比對出 {len(df_result)} 筆商品資料。")
                                st.dataframe(df_result, use_container_width=True)
                                
                                towrite_query = io.BytesIO()
                                with pd.ExcelWriter(towrite_query, engine='openpyxl') as writer:
                                    df_result.to_excel(writer, index=False, sheet_name="SKU查詢結果")
                                st.download_button("📥 下載本次SKU查詢結果報表 (.xlsx)", data=towrite_query.getvalue(), file_name=f"SKU批次查詢結果_{datetime.date.today().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                elif search_mode == "模糊關鍵字搜尋":
                    search_term = st.text_input("🔍 快速搜尋關鍵字 (支援條碼、品名、貨號模糊比對)：", placeholder="輸入搜尋內容...")
                    st.metric(label=f"📊 【{selected_sheet}】當前總資料筆數", value=f"{len(df_view)} 筆")
                    if search_term:
                        search_mask = df_view.astype(str).apply(lambda x: x.str.contains(search_term, case=False, na=False)).any(axis=1)
                        st.dataframe(df_view[search_mask], use_container_width=True)
                    else:
                        st.dataframe(df_view, use_container_width=True)
                        
            except Exception as e:
                st.error(f"❌ 讀取分頁數據失敗: {str(e)}")

    elif sub_page == "⚖️ 麗嬰商品表合併和與審核":
        st.subheader("🧸 麗嬰採購單一鍵導入與審核系統")
        try:
            if not cfg['ID_MASTER_FILE']:
                st.error("❌ 錯誤：未設定雲端主資料庫 ID (ID_MASTER_FILE)！")
                st.stop()
            master_data_result = load_master_data(cfg['ID_MASTER_FILE'])
            if master_data_result is None or not isinstance(master_data_result, (tuple, list)):
                st.error("❌ 讀取雲端主資料庫失敗。")
                st.stop()
            df_total, df_history, df_delete_log, df_meta, _, current_max_uid = master_data_result
            if df_total is None:
                st.error("❌ 讀取失敗：主資料庫內容為空。")
                st.stop()
        except Exception as e:
            st.error(f"❌ 發生例外錯誤：{str(e)}")
            st.stop()

        if "狀態" not in df_history.columns:
            df_history["狀態"] = ""

        if 'merge_success_msg' in st.session_state:
            st.success(st.session_state['merge_success_msg'])
            st.markdown("### ⚡ 歸檔後後續自動化推薦操作")
            if st.button("🚀 三表資料整合並自動回寫更新至雲端『商品蝦皮麗嬰價格統整表』", type="primary", use_container_width=True):
                with st.spinner("⏳ 正在回寫雲端..."):
                    if run_powerquery_and_update_gdrive():
                        st.success("🎯 狂賀！同步覆寫更新完畢！")
                        del st.session_state['merge_success_msg']   
            st.write("---")

        uploaded_files = st.file_uploader("📥 選擇採購單 Excel (可多選批次上傳)", type=["xlsx", "xls", "xlsm"], accept_multiple_files=True, key="main_merge_files")
        
        if uploaded_files:
            if st.button("🚀 開始一鍵合併到麗嬰總表", type="primary"):
                success_count = dup_count = no_barcode_count = anomaly_count = 0
                new_rows, history_records = [], []
                valid_dfs_to_merge = []
                
                master_dict = {}
                for idx, row in df_total.iterrows():
                    b_key = clean_barcode(row.get('條碼', ''))
                    if b_key and b_key != "0":
                        master_dict[b_key] = row
                
                history_md5_list = df_history['md5'].astype(str).tolist() if 'md5' in df_history.columns else []

                with st.spinner("⏳ 正在執行：智慧表頭解析、防錯清洗與狀態紀錄..."):
                    for file in uploaded_files:
                        file_bytes = file.read()
                        file_md5 = calculate_md5(file_bytes)
                        filename = file.name
                        now_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        
                        if file_md5 in history_md5_list:
                            dup_count += 1
                            history_records.append({"檔案名稱": filename, "md5": file_md5, "匯入時間": now_time, "狀態": "[重複檔案]"})
                            continue
                            
                        try:
                            df_src_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, nrows=15)
                            header_row = 0
                            for idx, row in df_src_raw.iterrows():
                                if row.astype(str).str.contains("條碼|國際條碼|EAN|Barcode|BARCODE").any():
                                    header_row = idx
                                    break
                            
                            smart_columns = process_smart_headers(df_src_raw, header_row)
                            df_src = pd.read_excel(io.BytesIO(file_bytes), skiprows=header_row + 1, header=None)
                            if df_src.empty or len(smart_columns) != len(df_src.columns):
                                df_src = pd.read_excel(io.BytesIO(file_bytes), header=header_row)
                            else:
                                df_src.columns = smart_columns

                            rename_dict = {}
                            for col in df_src.columns:
                                col_clean = str(col).strip()
                                if col_clean in ["名稱", "品名", "商品名稱", "中文", "中文品名", "品名規格", "Description"]: rename_dict[col] = "名稱"
                                elif col_clean in ["條碼", "國際條碼", "EAN", "Barcode", "BARCODE", "條碼型號", "JAN CODE"]: rename_dict[col] = "條碼"
                                elif col_clean in ["零售價", "建議售價", "單價", "定價", "售價", "Price"]: rename_dict[col] = "零售價"
                                elif col_clean in ["料品編號","商品編號", "貨號", "產品編號", "Item No", "ITEM"]: rename_dict[col] = "商品編號"
                                elif col_clean in ["內盒", "Inner", "INNER"]: rename_dict[col] = "內盒"
                                elif col_clean in ["CTN", "外箱", "箱入數", "Carton", "外箱數"]: rename_dict[col] = "CTN"
                                elif col_clean in ["CTN訂購含稅價", "CTN含稅價", "CTN含稅", "外箱含稅價", "外箱進價"]: rename_dict[col] = "CTN含稅"
                                elif col_clean in ["內盒訂購含稅價", "內盒含稅價", "內盒含稅", "內盒進價", "含稅", "含稅價", "進價"]: rename_dict[col] = "含稅"
                            df_src = df_src.rename(columns=rename_dict)
                            
                            if "條碼" not in df_src.columns:
                                no_barcode_count += 1
                                history_records.append({"檔案名稱": filename, "md5": file_md5, "匯入時間": now_time, "狀態": "[無條碼欄位]"})
                                continue
                                
                            df_src['條碼'] = df_src['條碼'].apply(clean_barcode)
                            df_valid = df_src[df_src['條碼'] != ""]
                            if df_valid.empty:
                                no_barcode_count += 1
                                history_records.append({"檔案名稱": filename, "md5": file_md5, "匯入時間": now_time, "狀態": "[無條碼欄位]"})
                                continue
                                
                            valid_dfs_to_merge.append({
                                "file_name": filename, "file_md5": file_md5, "df": df_valid
                            })
                            
                        except Exception as e:
                            st.error(f"❌ 檔案 【{filename}】 解析失敗: {str(e)}")
                            no_barcode_count += 1
                            history_records.append({"檔案名稱": filename, "md5": file_md5, "匯入時間": now_time, "狀態": "[無條碼欄位]"})
                            
                if valid_dfs_to_merge:
                    with st.spinner("⏳ 正在執行雲端 UID 流水號繼承與雙向交叉審核比對..."):
                        for item in valid_dfs_to_merge:
                            df_item = item["df"]
                            filename = item["file_name"]
                            file_md5 = item["file_md5"]
                            now_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                            for _, src_row in df_item.dropna(subset=['條碼']).iterrows():
                                barcode = str(src_row['條碼'])
                                if barcode in ["", "0"]: continue
                                
                                price = round(float(src_row.get('零售價', 0)), 2) if pd.notna(src_row.get('零售價', 0)) else 0
                                name = str(src_row.get('名稱', '')).strip()
                                
                                match_found = False
                                need_insert = False
                                remark_note = ""
                                display_barcode = barcode

                                if barcode in master_dict:
                                    match_found = True
                                    master_row = master_dict[barcode]
                                    target_uid = master_row.get('UID', '未知')
                                    m_name = str(master_row.get('名稱', '')).strip()
                                    try: m_price = float(master_row.get('零售價', 0)) if pd.notna(master_row.get('零售價', 0)) else 0
                                    except: m_price = 0
                                    
                                    if name and m_name != name: 
                                        need_insert = True
                                        anomaly_count += 1
                                        remark_note = f"與 UID: {target_uid} 條碼重複, 名稱不同"
                                        display_barcode = f"🔴 {barcode}"
                                    elif price > 0 and m_price != price: 
                                        need_insert = True
                                        anomaly_count += 1
                                        remark_note = f"與 UID: {target_uid} 條碼重複, 名稱相同, 零售價不同"
                                        display_barcode = f"🟢 {barcode}"
                                
                                if (not match_found) or (match_found and need_insert):
                                    current_max_uid += 1
                                    new_uid = f"UID-{int(current_max_uid):06d}"
                                    row_data = {col: "" for col in df_total.columns if col != 'move'}
                                    for col in df_item.columns:
                                        if col in row_data:
                                            val = src_row[col]
                                            row_data[col] = "" if pd.isna(val) else val
                                    
                                    row_data["UID"] = new_uid
                                    row_data["條碼"] = display_barcode
                                    row_data["名稱"] = name
                                    if "零售價" in row_data: row_data["零售價"] = price
                                    row_data["備註"] = remark_note
                                    row_data["匯入檔名"] = filename
                                    
                                    new_rows.append(row_data)
                                    master_dict[barcode] = row_data

                            success_count += 1
                            history_records.append({"檔案名稱": filename, "md5": file_md5, "匯入時間": now_time, "狀態": "[已匯入]"})
                            history_md5_list.append(file_md5)
                        
                if new_rows: df_total = pd.concat([df_total, pd.DataFrame(new_rows)], ignore_index=True)
                if history_records: df_history = pd.concat([df_history, pd.DataFrame(history_records)], ignore_index=True)
                
                df_meta.iloc[0, 0] = current_max_uid
                df_total = run_cross_matching(df_total)
                
                if save_to_master_xlsm({"麗嬰國際產品總表": df_total, "已處理採購單": df_history, "metadata": df_meta}, cfg['ID_MASTER_FILE'], cfg['ID_BASE_FOLDER'], cfg['NAME_MASTER']):
                    load_master_data.clear()
                    report_msg = f"🎉 成功完成狀態登記與資料同步！\n\n✅ [已匯入]: {success_count} 份\n🔁 [重複檔案]: {dup_count} 份\n⚠️ [無條碼欄位]: {no_barcode_count} 份"
                    if anomaly_count > 0:
                        report_msg += f"\n\n🚨 注意：本次匯入發現 **{anomaly_count}** 筆異常衝突商品！"
                    st.session_state['merge_success_msg'] = report_msg
                    st.rerun()

        st.write("---")
        st.subheader("⚠️ 條碼重複與衝突即時審核控制台")
        if not df_total.empty:
            if 'move' in df_total.columns: df_total = df_total.drop(columns=['move'])
            df_total.insert(0, 'move', False)
            df_total['move'] = df_total['move'].astype(bool)
            
            df_total['條碼_乾淨'] = df_total['條碼'].astype(str).str.replace('🔴 ', '').str.replace('🟢 ', '').str.strip()
            is_duplicate_barcode = df_total.duplicated(subset=['條碼_乾淨'], keep=False) & (~df_total['條碼_乾淨'].isin(["", "0", "nan", "None"]))
            
            df_anomaly = df_total[is_duplicate_barcode].sort_values(by=['條碼_乾淨', 'UID']).copy()
            df_anomaly = df_anomaly.drop(columns=['條碼_乾淨'], errors='ignore')
            df_total = df_total.drop(columns=['條碼_乾淨'], errors='ignore')
            
            if not df_anomaly.empty:
                def inject_emoji_alerts(row):
                    remark = str(row.get('備註', ''))
                    barcode = str(row.get('條碼', ''))
                    if "名稱不同" in remark and not barcode.startswith("🔴"): row['條碼'] = f"🔴 {barcode}"
                    elif "零售價不同" in remark and not barcode.startswith("🟢"): row['條碼'] = f"🟢 {barcode}"
                    return row
                df_anomaly = df_anomaly.apply(inject_emoji_alerts, axis=1)
                
                st.warning("下方商品為系統抓出之條碼重複資料：🔴 代表名稱不一致，🟢 代表售價不一致。")
                edited_anomaly_df = st.data_editor(df_anomaly, use_container_width=True, disabled=[col for col in df_anomaly.columns if col not in ['move', '備註']], key="anomaly_editor")
                
                if st.button("🧹 執行審核與資料儲存", type="primary"):
                    for index, edited_row in edited_anomaly_df.iterrows():
                        target_uid = edited_row['UID']
                        new_note = str(edited_row['備註']).strip().replace("🔴 ", "").replace("🟢 ", "")
                        df_total.loc[df_total['UID'] == target_uid, '備註'] = new_note

                    uids_to_delete = edited_anomaly_df[edited_anomaly_df['move'] == True]['UID'].values
                    if len(uids_to_delete) > 0:
                        df_to_delete = df_total[df_total['UID'].isin(uids_to_delete)].copy()
                        df_to_delete['刪除時間'] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        df_to_delete = df_to_delete.drop(columns=['move'], errors='ignore')
                        
                        df_remaining = df_total[~df_total['UID'].isin(uids_to_delete)].copy()
                        df_remaining = df_remaining.drop(columns=['move'], errors='ignore')
                        df_remaining = run_cross_matching(df_remaining)
                        df_delete_log = pd.concat([df_delete_log, df_to_delete], ignore_index=True)
                        
                        if save_to_master_xlsm({"麗嬰國際產品總表": df_remaining, "刪除紀錄": df_delete_log}, cfg['ID_MASTER_FILE'], cfg['ID_BASE_FOLDER'], cfg['NAME_MASTER']):
                            load_master_data.clear() 
                            st.session_state['merge_success_msg'] = "🧹 移轉封存與自訂 Note 備註已同步寫入！"
                            st.rerun()
                    else:
                        df_remaining_only = run_cross_matching(df_total.drop(columns=['move'], errors='ignore'))
                        if save_to_master_xlsm({"麗嬰國際產品總表": df_remaining_only, "刪除紀錄": df_delete_log}, cfg['ID_MASTER_FILE'], cfg['ID_BASE_FOLDER'], cfg['NAME_MASTER']):
                            load_master_data.clear() 
                            st.session_state['merge_success_msg'] = "📝 備註內容已同步更新！"
                            st.rerun()
            else:
                st.success("🟢 當前總表中沒有任何重複商品的衝突。")

    elif sub_page == "📈 蝦皮商品清單轉換":
        st.subheader("🛍️ 蝦皮賣場商品列表iSKU結構校正")
        df_shopee_history, df_shopee_current_list = load_shopee_data(cfg['ID_SHOPEE_MASTER'])

        uploaded_shopee = st.file_uploader("📥 上傳新的蝦皮商品清單原始報表：", type=["xlsx", "xls", "xlsm"], key="main_shopee_upload")
        if uploaded_shopee:
            file_bytes = uploaded_shopee.read()
            shopee_md5 = calculate_md5(file_bytes)
            
            if shopee_md5 in df_shopee_history['md5'].astype(str).values:
                st.error(f"⚠️ 拒絕重複格式校正！系統已自動封鎖。")
            else:
                if st.button("🪄 執行蝦皮iSKU結構校正", type="primary", use_container_width=True):
                    try:
                        df_shopee_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='openpyxl')
                        if df_shopee_raw.shape[1] >= 11: df_shopee_raw.drop(df_shopee_raw.columns[10], axis=1, inplace=True)
                        shopee_headers = df_shopee_raw.iloc[2].astype(str).str.strip().tolist()
                        df_shopee = df_shopee_raw.iloc[6:].copy()
                        df_shopee.columns = shopee_headers
                        df_shopee.reset_index(drop=True, inplace=True)
                        
                        def calc_isku_row(row):
                            opt = str(row.get('商品選項貨號', '')).strip()
                            main = str(row.get('主商品貨號', '')).strip()
                            if opt in ["見選項", "null", "Null", "nan", "NaN", "None"]: opt = ""
                            if main in ["見選項", "null", "Null", "nan", "NaN", "None"]: main = ""
                            return opt if opt != "" else (main if main != "" else "蝦皮無iSKU")
                            
                        df_shopee['iSKU'] = df_shopee.apply(calc_isku_row, axis=1)
                        df_shopee['original_index'] = df_shopee.index
                        cols_list = list(df_shopee.columns)
                        if "iSKU" in cols_list and "價格" in cols_list:
                            cols_list.remove("iSKU")
                            cols_list.insert(cols_list.index("價格"), "iSKU")
                            df_shopee = df_shopee[cols_list]
                        
                        df_valid_isku = df_shopee[df_shopee['iSKU'] != "蝦皮無iSKU"].copy()
                        df_isku_keep = df_valid_isku.sort_values(by=['iSKU', '價格', 'original_index']).drop_duplicates(subset=['iSKU'], keep='last')
                        df_gtin_check = df_isku_keep.copy()
                        df_gtin_check['GTIN_str'] = df_gtin_check['GTIN'].astype(str).str.strip().str.split('.').str[0]
                        df_gtin_keep = df_gtin_check[~df_gtin_check['GTIN_str'].isin(["", "00", "0", "nan"])].sort_values(by=['GTIN_str', '價格', 'original_index']).drop_duplicates(subset=['GTIN_str'], keep='last')
                        df_final_clean = pd.concat([df_gtin_keep, df_gtin_check[df_gtin_check['GTIN_str'].isin(["", "00", "0", "nan"])]]).sort_values(by='original_index')
                                         
                        new_hist_log = pd.DataFrame([{"檔案名稱": uploaded_shopee.name, "md5": shopee_md5, "匯入時間": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}])
                        df_shopee_history = pd.concat([df_shopee_history, new_hist_log], ignore_index=True)
                        
                        if save_to_shopee_master_xlsm({"蝦皮商品列表": df_final_clean, "匯入檔案": df_shopee_history}, cfg['ID_SHOPEE_MASTER'], cfg['ID_SHOPEE_FOLDER'], cfg['NAME_SHOPEE']):
                            load_shopee_data.clear() 
                            get_cached_gdrive_id.clear()
                            st.session_state['shopee_clean'] = df_final_clean
                            st.success(f"🎉 校正完成！已覆寫雲端。")
                    except Exception as e:
                        st.error(f"讀取或清洗蝦皮檔案失敗: {str(e)}")

        if 'shopee_clean' in st.session_state:
            st.dataframe(st.session_state['shopee_clean'], use_container_width=True)
            towrite_shopee = io.BytesIO()
            st.session_state['shopee_clean'].to_excel(towrite_shopee, index=False)
            st.download_button(label="📥 下載此次iSKU校正蝦皮報表 (.xlsx)", data=towrite_shopee.getvalue(), file_name=f"蝦皮清洗完成對齊表_{datetime.date.today().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    elif sub_page == "🔍 商品清單紀錄查詢":
        st.subheader("📊 歷史商品清單紀錄查詢")
        hist_files = list_gdrive_files(cfg['ID_PROD_FOLDER'])
        if not hist_files: 
            st.warning(f"💡 目前雲端無歷史單據。")
        else:
            file_options = {f['name']: f['id'] for f in hist_files}
            selected_hist_file = st.selectbox("🎯 選擇商品清單紀錄檔案：", list(file_options.keys()))
            
            if selected_hist_file:
                try:
                    target_id = file_options[selected_hist_file]
                    file_bytes = get_cached_gdrive_file_bytes(target_id)
                    engine_kw = {"engine": "calamine"} if HAS_CALAMINE else {}
                    df_hist_view = pd.read_excel(io.BytesIO(file_bytes), **engine_kw)
                    
                    st.markdown(f"📄 **當前雲端檔案**：`{selected_hist_file}` ｜ 📊 **總品項數**：`{len(df_hist_view)} 筆`")
                    st.write("---")
                    
                    st.markdown("#### 🔎 多筆批次編碼查詢")
                    col_m, col_i = st.columns([1, 3])
                    with col_m:
                        target_col = st.radio("選擇查詢依據欄位：", options=["自定義編碼", "c"], index=0)
                    with col_i:
                        batch_input = st.text_area(f"請輸入多筆【{target_col}】（每筆請以換行、逗號或空格隔開）：", height=100)
                    
                    df_display = df_hist_view.copy()
                    if batch_input.strip() and target_col in df_display.columns:
                        import re
                        search_terms = [t.strip() for t in re.split(r'[\n,\s]+', batch_input) if t.strip()]
                        df_display[target_col] = df_display[target_col].astype(str).str.strip().str.split('.').str[0]
                        df_display = df_display[df_display[target_col].isin(search_terms)]
                        st.info(f"🎯 批次篩選結果：找到 **{len(df_display)}** 筆符合資料。")
                    
                    st.dataframe(df_display, use_container_width=True)
                    
                    towrite_prod = io.BytesIO()
                    with pd.ExcelWriter(towrite_prod, engine='openpyxl') as writer:
                        df_display.to_excel(writer, index=False, sheet_name="商品清單查詢結果")
                    
                    st.download_button(
                        label="🔄 下載此歷史商品清單 (或篩選結果)", 
                        data=towrite_prod.getvalue(), 
                        file_name=selected_hist_file, 
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e: 
                    st.error(f"❌ 讀取失敗: {str(e)}")