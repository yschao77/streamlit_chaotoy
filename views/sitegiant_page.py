import sys
import os

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import streamlit as st
import pandas as pd
import io
import datetime

# 從外層的 utils.py 引入我們需要的工具
from utils import (
    HAS_CALAMINE, 
    download_gdrive_file_to_bytes, 
    upload_or_update_gdrive_file,
    list_gdrive_files,
    get_cached_gdrive_file_bytes,
    clean_barcode,
    load_shopee_data
)

def render(sub_page, ID_PRICE_SUMMARY, ID_HISTORY_INWARD_FOLDER, ID_SHOPEE_MASTER):
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
                                        
                                        shopee_name = match_row.get('蝦皮商品名稱', None)
                                        internal_name = match_row.get('內部商品名稱', match_row.get('名稱', None)) 
                                        
                                        if pd.notna(shopee_name) and str(shopee_name).strip() not in ["", "nan", "None"]:
                                            prod_name = str(shopee_name).strip()
                                        elif pd.notna(internal_name) and str(internal_name).strip() not in ["", "nan", "None"]:
                                            prod_name = str(internal_name).strip()
                                        
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
                                
                            if missing_items:
                                try:
                                    TARGET_SHEET_ID = "1Ixp9V_u2yU8hiWhxQCHNDB4kPKlxDGD2"
                                    try:
                                        raw_bytes = download_gdrive_file_to_bytes(TARGET_SHEET_ID)
                                        engine_kw = {"engine": "calamine"} if HAS_CALAMINE else {}
                                        df_missing = pd.read_excel(raw_bytes, sheet_name=0, dtype=str, **engine_kw)
                                    except Exception:
                                        df_missing = pd.DataFrame(columns=["採購單檔名", "國際條碼", "狀況", "狀態", "建立時間"])
                                        
                                    df_missing.columns = df_missing.columns.astype(str).str.strip()
                                    
                                    if "採購單檔名" in df_missing.columns and "國際條碼" in df_missing.columns:
                                        existing_keys = set(
                                            df_missing["採購單檔名"].astype(str).str.strip() + "_" + 
                                            df_missing["國際條碼"].astype(str).str.strip()
                                        )
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
                                        st.toast(f"🚨 已自動將 {len(new_missing_items)} 筆異常紀錄向下新增至雲端！", icon="⚠️")
                                except Exception as log_err:
                                    st.error(f"⚠️ 自動記錄異常商品失敗: {str(log_err)}")
                                    
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
        st.subheader("📊 歷史入庫單與成本稅款加總指標檢視")
        hist_files = list_gdrive_files(ID_HISTORY_INWARD_FOLDER)
        if not hist_files: st.warning(f"💡 目前雲端無歷史單據。")
        else:
            file_options = {f['name']: f['id'] for f in hist_files}
            selected_hist_file = st.selectbox("🎯 選擇欲調閱的入庫對帳單：", list(file_options.keys()))
            if selected_hist_file:
                try:
                    target_id = file_options[selected_hist_file]
                    file_bytes = download_gdrive_file_to_bytes(target_id)
                    df_hist_view = pd.read_excel(file_bytes, engine="calamine" if HAS_CALAMINE else None)
                    st.markdown(f"📄 **當前雲端檔案**：`{selected_hist_file}` ｜ 📊 **單據品項數**：`{len(df_hist_view)} 筆`")
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
                    
                    st.download_button(label="🔄 下載此歷史採購入庫單", data=file_bytes.getvalue(), file_name=selected_hist_file, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                except Exception as e: st.error(f"❌ 讀取失敗: {str(e)}")

    # -------------------------------------------------------------------------
    # 子功能 3：📦 SiteGiant 批量新增UPC
    # -------------------------------------------------------------------------
    elif sub_page == "📦 Sitegiant 批量新增UPC":
        st.subheader("📦 Sitegiant 批量新增UPC")
        
        with st.expander("📌 點擊查看操作流程與前/後置條件", expanded=True):
            st.markdown("""
            ### 📝 前置條件
            1. **確認 iSKU 對應 UPC**。
            2. **下載 Sitegiant 批量編輯 UPC 檔案**：
               - 至 Sitegiant 庫存列表 ➔ [批量編輯](https://sitegiant.co/items/batch-edit) UPC ➔ 下載 Excel。
            3. **確認蝦皮賣場列表已校正**。
            ### 🚀 後續操作
            - 將下方處理完畢並下載的新 Excel 檔案，上傳回 Sitegiant 覆蓋即可。
            """)

        df_shopee_hist, df_shopee_list = load_shopee_data(ID_SHOPEE_MASTER)

        if df_shopee_list.empty:
            st.error("❌ 無法從雲端讀取蝦皮商品列表！請先前往「蝦皮商品清單轉換」執行校正並回寫雲端。")
        else:
            st.info("✅ 系統已自動從雲端載入最新的蝦皮商品列表，準備好進行 UPC 交叉比對！")
            
            uploaded_sg_file = st.file_uploader("📥 請上傳 Sitegiant UPC 批量編輯下載檔 (.xlsx/.xls/.csv)", type=["xlsx", "xls", "csv"])
            
            if uploaded_sg_file:
                if st.button("⚡ 執行自動比對並填補 UPC", type="primary", use_container_width=True):
                    with st.spinner("⏳ 正在自動比對蝦皮資料庫並填入缺失的 UPC..."):
                        try:
                            if uploaded_sg_file.name.lower().endswith('.csv'):
                                df_sg = pd.read_csv(uploaded_sg_file, dtype=str)
                            else:
                                df_sg = pd.read_excel(uploaded_sg_file, dtype=str)
                            
                            df_sg.columns = df_sg.columns.astype(str).str.strip().str.replace('\n', '')
                            
                            sg_sku_col = next((c for c in ["Item SKU", "商品 SKU", "SKU", "Item Sku", "item sku", "庫存SKU"] if c in df_sg.columns), None)
                            sg_upc_col = next((c for c in ["UPC", "國際條碼（UPC）", "國際條碼", "upc"] if c in df_sg.columns), None)
                            sg_main_col = next((c for c in ["Is Main UPC", "主要"] if c in df_sg.columns), None)
                            
                            if not sg_sku_col or not sg_upc_col:
                                st.error("❌ 檔案解析失敗：找不到對應的 SKU 或 UPC 欄位。")
                            else:
                                df_shopee_list['iSKU'] = df_shopee_list['iSKU'].astype(str).str.strip()
                                df_shopee_list['GTIN_str'] = df_shopee_list['GTIN'].astype(str).str.strip().str.split('.').str[0]
                                valid_shopee = df_shopee_list[~df_shopee_list['GTIN_str'].isin(["", "00", "0", "nan", "#N/A", "None", "空白"])]
                                
                                upc_map_exact = dict(zip(valid_shopee['iSKU'], valid_shopee['GTIN_str']))
                                                        
                                updated_indices = []
                                
                                for idx, row in df_sg.iterrows():
                                    sku = str(row[sg_sku_col]).strip()
                                    current_upc = clean_barcode(row.get(sg_upc_col, ""))
                                    
                                    if current_upc in ["", "nan", "None", "0", "00"]:
                                        if sku in upc_map_exact:
                                            match_gtin = upc_map_exact[sku]
                                            df_sg.at[idx, sg_upc_col] = match_gtin
                                            if sg_main_col:
                                                df_sg.at[idx, sg_main_col] = "是" if "主要" in sg_main_col else "Yes"
                                            updated_indices.append(idx)
                                
                                df_sg_filtered = df_sg.loc[updated_indices].reset_index(drop=True)
                                st.session_state['sg_upc_updated_df'] = df_sg_filtered
                                
                                if len(df_sg_filtered) > 0:
                                    st.success(f"🎉 處理完成！共成功自動填補 **{len(df_sg_filtered)}** 筆缺失的 UPC 資料。")
                                else:
                                    st.warning("⚠️ 處理完成，但未找到任何可填補的缺失 UPC 資料。")
                                
                        except Exception as e:
                            st.error(f"❌ 處理檔案時發生錯誤：{str(e)}")

            if 'sg_upc_updated_df' in st.session_state and not st.session_state['sg_upc_updated_df'].empty:
                df_result = st.session_state['sg_upc_updated_df']
                st.markdown(f"### 📋 成功新增 UPC 預覽（共 {len(df_result)} 筆）")
                st.dataframe(df_result, use_container_width=True)
                
                towrite_sg = io.BytesIO()
                with pd.ExcelWriter(towrite_sg, engine='openpyxl') as writer:
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