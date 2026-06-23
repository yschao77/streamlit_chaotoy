import streamlit as st
import pandas as pd
import openpyxl
import hashlib
import datetime
import io
import os
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

# =========================================================================
# 🛠️ 🔴 全域攔截並修補 openpyxl 描述器核心驗證 Bug (Monkey Patch)
# =========================================================================
try:
    import openpyxl.descriptors.base
    orig_set_attr = openpyxl.descriptors.base.Set.__set__

    def patched_set_attr(self, instance, value):
        if isinstance(value, str) and '-' in value:
            parts = value.split('-')
            value = parts[0] + ''.join(p.title() for p in parts[1:])
        try:
            orig_set_attr(self, instance, value)
        except ValueError:
            if hasattr(self, 'values'):
                if isinstance(self.values, set):
                    self.values.add(value)
                elif isinstance(self.values, tuple):
                    self.values = self.values + (value,)
                elif isinstance(self.values, list):
                    self.values.append(value)
            orig_set_attr(self, instance, value)
    openpyxl.descriptors.base.Set.__set__ = patched_set_attr
except Exception:
    pass

# =========================================================================
# 🌐 1. Google Drive API 安全連線初始化與雲端檔案 ID 配置
# =========================================================================
st.set_page_config(page_title="麗嬰採購與入庫雲端整合系統", layout="wide")

# 由 Streamlit 側邊欄動態輸入雲端實體 ID（免硬編碼，彈性最高）
st.sidebar.header("🌐 Google Drive 雲端連線設定")
MASTER_FILE_ID = st.sidebar.text_input("1. 核心產品總表 (xlsm) File ID:", help="輸入該 Excel 雲端網址中 d/ 後方的那串亂碼")
PRICE_SUMMARY_ID = st.sidebar.text_input("2. 商品價格統整表 File ID:", help="輸入統整表網址中 d/ 後方的那串亂碼")
HISTORY_FOLDER_ID = st.sidebar.text_input("3. 歷史入庫單備份夾 Folder ID:", help="輸入雲端資料夾網址中 folders/ 後方的那串亂碼")

@st.cache_resource
def init_drive_service():
    # 讀取部署後設定在 Streamlit Secrets 的金鑰字典
    google_secrets = st.secrets["textkey"]
    credentials = Credentials.from_service_account_info(
        google_secrets,
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build('drive', 'v3', credentials=credentials)

# 驗證輸入並初始化 API
if not (MASTER_FILE_ID and PRICE_SUMMARY_ID and HISTORY_FOLDER_ID):
    st.warning("💡 請先在左側控制台輸入完整的 Google Drive 檔案與資料夾 ID 以啟用雲端資料庫。")
    st.stop()

try:
    drive_service = init_drive_service()
except Exception as e:
    st.error(f"❌ Google Drive API 連線失敗，請檢查憑證。錯誤資訊: {str(e)}")
    st.stop()

# =========================================================================
# 📦 2. 雲端專用讀寫與核心工具函式
# =========================================================================
def calculate_md5(file_bytes):
    return hashlib.md5(file_bytes).hexdigest()

def download_file_from_drive(file_id):
    """將雲端檔案下載至記憶體，避免產生伺服器實體快取"""
    request = drive_service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.seek(0)
    return fh

def save_to_master_xlsm_cloud(sheets_dict, file_id):
    """直接將修改後的資料回寫並覆蓋 Google Drive 上的主 xlsm 檔案 (保留 VBA)"""
    try:
        # 1. 先從雲端把原本含有巨集巨集的原始檔案載入記憶體
        origin_fh = download_file_from_drive(file_id)
        wb = openpyxl.load_workbook(origin_fh, keep_vba=True)
        
        # 2. 刷新對應的工作表資料
        for sheet_name, df in sheets_dict.items():
            if sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                ws.delete_rows(1, ws.max_row + 1)
            else:
                ws = wb.create_sheet(sheet_name)
            ws.append(list(df.columns))
            
            barcode_col_idx = list(df.columns).index("條碼") + 1 if "條碼" in df.columns else None
            for row_idx, row in enumerate(df.itertuples(index=False), start=2):
                cleaned_row = []
                for col_idx, x in enumerate(row):
                    if pd.isna(x): cleaned_row.append("")
                    elif barcode_col_idx and (col_idx + 1) == barcode_col_idx:
                        cleaned_row.append(str(x).strip().split('.')[0])
                    else: cleaned_row.append(x)
                ws.append(cleaned_row)
                if barcode_col_idx:
                    ws.cell(row=row_idx, column=barcode_col_idx).number_format = '@'
                    
        # 3. 儲存回記憶體串流
        output_fh = io.BytesIO()
        wb.save(output_fh)
        output_fh.seek(0)
        
        # 4. 透過 API 覆蓋上傳回原檔案
        media = MediaFileUpload(output_fh, mimetype='application/vnd.ms-excel.sheet.macroEnabled.12', resumable=True)
        drive_service.files().update(fileId=file_id, media_body=media).execute()
        return True
    except Exception as e:
        st.error(f"❌ 寫入雲端主資料庫發生錯誤: {str(e)}")
        return False

def run_cross_matching(df):
    if df.empty: return df
    df['備註'] = df['備註'].astype(str).apply(lambda x: "" if "條碼重複" in x or x == "nan" else x)
    records = df.to_dict('records')
    from collections import defaultdict
    barcode_groups = defaultdict(list)
    for idx, row in enumerate(records):
        b_str = str(row.get('條碼', '')).strip().split('.')[0]
        if b_str and b_str not in ["", "0", "nan", "None"]:
            barcode_groups[b_str].append(idx)
            
    for b_str, indices in barcode_groups.items():
        if len(indices) > 1:
            for check_idx in indices:
                check_row = records[check_idx]
                name_str = str(check_row.get('名稱', '')).strip()
                try: price_val = float(check_row.get('零售價', 0)) if pd.notna(check_row.get('零售價', 0)) else 0
                except: price_val = 0
                for comp_idx in indices:
                    if check_idx == comp_idx: continue
                    comp_row = records[comp_idx]
                    comp_name_str = str(comp_row.get('名稱', '')).strip()
                    try: comp_price_val = float(comp_row.get('零售價', 0)) if pd.notna(comp_row.get('零售價', 0)) else 0
                    except: comp_price_val = 0
                    uid_str = str(comp_row.get('UID', '未知')).strip()
                    if name_str != comp_name_str:
                        records[check_idx]['備註'] = f"與 UID: {uid_str} 條碼重複, 名稱不同"
                        break
                    elif price_val != comp_price_val:
                        records[check_idx]['備註'] = f"與 UID: {uid_str} 條碼重複, 名稱相同, 零售價不同"
                        break
    return pd.DataFrame(records)

# =========================================================================
# ⚙️ 3. 雲端資料庫預先連線載入
# =========================================================================
try:
    master_fh = download_file_from_drive(MASTER_FILE_ID)
    with pd.ExcelFile(master_fh) as xls:
        df_total = pd.read_excel(xls, "麗嬰國際產品總表")
        df_history = pd.read_excel(xls, "已匯入採購單")
        df_delete_log = pd.read_excel(xls, "刪除紀錄") if "刪除紀錄" in xls.sheet_names else pd.DataFrame(columns=["UID", "名稱", "條碼", "零售價", "備註", "匯入檔名", "刪除時間"])
        df_meta = pd.read_excel(xls, "metadata")
        all_sheets = xls.sheet_names
        
    if "條碼" in df_total.columns:
        df_total['條碼'] = df_total['條碼'].astype(str).str.strip().str.split('.').str[0]
    if "條碼" in df_delete_log.columns:
        df_delete_log['條碼'] = df_delete_log['條碼'].astype(str).str.strip().str.split('.').str[0]
    if not df_history.empty:
        df_history.columns = [str(col).strip().lower() for col in df_history.columns]
        df_history = df_history.loc[:, ~df_history.columns.duplicated()].copy()
        standard_cols = ["檔案名稱", "md5", "匯入時間"]
        if len(df_history.columns) >= 3: df_history.columns = standard_cols + list(df_history.columns[3:])
        else: df_history = pd.DataFrame(columns=standard_cols)
    else:
        df_history = pd.DataFrame(columns=["檔案名稱", "md5", "匯入時間"])
    current_max_uid = int(df_meta.iloc[0, 0]) if not df_meta.empty else 3473
except Exception as e:
    st.error(f"❌ 讀取雲端主資料庫失敗: {str(e)}。請確認該檔案已共用給服務帳戶，且 ID 輸入正確。")
    st.stop()

if 'inward_input_df' not in st.session_state:
    st.session_state['inward_input_df'] = pd.DataFrame([{"國際條碼": "", "數量": 1}])

# =========================================================================
# 🗂️ 4. 全域左側功能表單控制台 (Sidebar 面板)
# =========================================================================
with st.sidebar:
    st.write("---")
    st.header("🎯 請選擇功能模組：")
    menu = st.radio("功能切換：", ["🧸 麗嬰採購單合併與審核", "🔍 麗嬰總表分頁檢視", "🧡 蝦皮商品列表優化", "📊 PowerQuery 三表整合", "🛍️ 採購入庫單模組"])
    st.write("---")
    
    if menu == "🧸 麗嬰採購單合併與審核":
        st.subheader("📥 採購單批次上傳")
        uploaded_files = st.file_uploader("選擇採購單 Excel", type=["xlsx", "xls", "xlsm"], accept_multiple_files=True, key="side_files")
        btn_merge = st.button("🚀 開始一鍵合併歸檔", type="primary", use_container_width=True)
    elif menu == "🔍 麗嬰總表分頁檢視":
        st.subheader("👁️ 檢視與搜尋設定")
        view_sheets = [s for s in all_sheets if s != "麗嬰產品新採購單"]
        selected_sheet = st.selectbox("請選擇數據分頁：", view_sheets)
        search_term = st.text_input("🔍 快速搜尋關鍵字：", placeholder="輸入搜尋內容...")
    elif menu == "🧡 蝦皮商品列表優化":
        st.subheader("📥 蝦皮報表優化")
        uploaded_shopee = st.file_uploader("選擇蝦皮商品 Excel 檔案", type=["xlsx", "xls"], key="side_shopee")
        btn_shopee = st.button("🪄 執行蝦皮大清洗", type="primary", use_container_width=True)
    elif menu == "📊 PowerQuery 三表整合":
        st.subheader("🛠️ 本地聯結輔助檔案 ID 輸入")
        LOCAL_PROD_ID = st.text_input("商品iSKU清單檔案 ID:", help="對接 [商品列表.xlsx]")
        LOCAL_SHOPEE_ID = st.text_input("蝦皮商品列表檔案 ID:", help="對接 [蝦皮商品清單.xlsx]")
        btn_pq = st.button("🛠️ 啟動智慧勾稽整合", type="primary", use_container_width=True)
    elif menu == "🛍️ 採購入庫單模組":  
        st.subheader("📦 入庫單基礎設定")
        order_no = st.text_input("📝 請輸入訂單/銷貨單號：", value=datetime.date.today().strftime("%Y%m%d01"))
        vendor_name = st.selectbox("🏬 請選擇廠商清單：", ["麗嬰", "buyee", "日亞"])
        recv_date = st.date_input("📅 選擇收貨日：", value=datetime.date.today())
        
        st.write("---")
        st.subheader("🚀 條碼批次快速貼上區")
        bulk_paste_area = st.text_area(
            "請在此處貼上條碼明細：", 
            height=180, 
            placeholder="支援格式一（直接複製Excel）：\n4904810486527\t5\n4904810209638\t10\n\n支援格式二（純條碼）：\n4904810913207",
        )
        btn_inward = st.button("✨ 執行入庫單轉換", type="primary", use_container_width=True)

# =========================================================================
# 💻 5. 右側情報數據大畫面 (Main Area)
# =========================================================================
if menu == "🧸 麗嬰採購單合併與審核":
    st.header("🧸 麗嬰採購單雲端導入與審核系統")
    if uploaded_files and btn_merge:
        success_count = 0; dup_count = 0; no_barcode_count = 0
        new_rows = []; history_records = []
        
        master_dict = {str(row['條碼']).strip(): row for _, row in df_total.iterrows() if pd.notna(row['條碼'])}
        for file in uploaded_files:
            file_bytes = file.read()
            file_md5 = calculate_md5(file_bytes)
            if file_md5 in df_history['md5'].astype(str).values:
                dup_count += 1
                continue
            try:
                df_src_raw = pd.read_excel(io.BytesIO(file_bytes), header=None)
                header_row = 0
                for idx, row in df_src_raw.iterrows():
                    if row.astype(str).str.contains("條碼|國際條碼|EAN|Barcode|品名|名稱|零售價|單價").any():
                        header_row = idx
                        break
                df_src = pd.read_excel(io.BytesIO(file_bytes), header=header_row)
                df_src.columns = [str(col).strip() if ("Unnamed:" not in str(col) and str(col) != "nan") else "" for col in df_src.columns]
                
                rename_dict = {}
                for col in df_src.columns:
                    if col == "": continue
                    col_clean = str(col).strip()
                    if col_clean in ["名稱", "品名", "商品名稱", "中文", "中文品名", "品名規格", "Description"]: rename_dict[col] = "名稱"
                    elif col_clean in ["條碼", "國際條碼", "EAN", "Barcode", "BARCODE", "條碼型號", "JAN CODE"]: rename_dict[col] = "條碼"
                    elif col_clean in ["零售價", "建議售價", "單價", "定價", "售價", "Price"]: rename_dict[col] = "零售價"
                    elif col_clean in ["商品編號", "貨號", "產品編號", "Item No", "ITEM"]: rename_dict[col] = "商品編號"
                    elif col_clean in ["內盒", "Inner", "INNER"]: rename_dict[col] = "內盒"
                    elif col_clean in ["CTN", "外箱", "箱入數", "Carton", "外箱數"]: rename_dict[col] = "CTN"
                    elif col_clean in ["CTN訂購含稅價", "CTN含稅價", "CTN含稅", "外箱含稅價", "外箱進價"]: rename_dict[col] = "CTN含稅"
                    elif col_clean in ["內盒訂購含稅價", "內盒含稅價", "內盒含稅", "內盒進價", "含稅", "含稅價", "進價"]: rename_dict[col] = "含稅"
                df_src = df_src.rename(columns=rename_dict)
                if "條碼" not in df_src.columns:
                    no_barcode_count += 1
                    continue
                df_src['條碼'] = df_src['條碼'].astype(str).str.strip().str.split('.').str[0]
                
                for _, src_row in df_src.dropna(subset=['條碼']).iterrows():
                    barcode = str(src_row['條碼'])
                    if barcode in ["", "0", "nan", "None"]: continue
                    price = round(float(src_row.get('零售價', 0)), 2) if pd.notna(src_row.get('零售價', 0)) else 0
                    name = str(src_row.get('名稱', '')).strip()
                    match_found = False; need_insert = False
                    if barcode in master_dict:
                        match_found = True
                        master_row = master_dict[barcode]
                        if name and str(master_row['名稱']).strip() != name: need_insert = True
                        else:
                            try: m_price = float(master_row['零售價']) if (pd.notna(master_row['零售價']) and str(master_row['零售價']).strip() != "") else 0
                            except: m_price = 0
                            if price > 0 and m_price != price: need_insert = True
                    if (not match_found) or (match_found and need_insert):
                        current_max_uid += 1
                        new_uid = f"UID-{current_max_uid:06d}"
                        row_data = {col: "" for col in df_total.columns if col != 'move'}
                        for col in df_src.columns:
                            if col in row_data:
                                val = src_row[col]
                                row_data[col] = "" if pd.isna(val) else val
                        row_data["UID"] = new_uid
                        row_data["條碼"] = barcode
                        row_data["名稱"] = name
                        if "零售價" in row_data: row_data["零售價"] = price
                        row_data["備註"] = ""
                        row_data["匯入檔名"] = file.name
                        new_rows.append(row_data)
                
                history_records.append({
                    "檔案名稱": file.name, "md5": file_md5, 
                    "匯入時間": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                })
                success_count += 1
            except Exception as e:
                st.error(f"❌ 解析失敗: {str(e)}")
                
        if new_rows: df_total = pd.concat([df_total, pd.DataFrame(new_rows)], ignore_index=True)
        if history_records: df_history = pd.concat([df_history, pd.DataFrame(history_records)], ignore_index=True)
        df_meta.iloc[0, 0] = current_max_uid
        df_total = run_cross_matching(df_total)
        
        if save_to_master_xlsm_cloud({"麗嬰國際產品總表": df_total, "已匯入採購單": df_history, "metadata": df_meta}, MASTER_FILE_ID):
            st.success("🎉 雲端資料庫同步完成！")
            st.rerun()

    st.subheader("2. 異常條碼即時審核控制台")
    if not df_total.empty:
        if 'move' in df_total.columns: df_total = df_total.drop(columns=['move'])
        df_total.insert(0, 'move', False)
        df_total['move'] = df_total['move'].astype(bool)
        is_duplicate_barcode = df_total.duplicated(subset=['條碼'], keep=False) & (~df_total['條碼'].isin(["", "0", "nan", "None"]))
        df_anomaly = df_total[is_duplicate_barcode].sort_values(by=['條碼', 'UID']).copy()
        if not df_anomaly.empty:
            def inject_emoji_alerts(row):
                remark = str(row.get('備註', ''))
                barcode = str(row.get('條碼', ''))
                if "名稱不同" in remark and not barcode.startswith("🔴"): row['條碼'] = f"🔴 {barcode}"
                elif "零售價不同" in remark and not barcode.startswith("🟢"): row['條碼'] = f"🟢 {barcode}"
                return row
            df_anomaly = df_anomaly.apply(inject_emoji_alerts, axis=1)
            edited_anomaly_df = st.data_editor(df_anomaly, use_container_width=True, disabled=[col for col in df_anomaly.columns if col != 'move'], key="anomaly_editor")
            if st.button("🧹 執行審核與刪除紀錄封存", type="primary"):
                uids_to_delete = edited_anomaly_df[edited_anomaly_df['move'] == True]['UID'].values
                if len(uids_to_delete) > 0:
                    df_to_delete = df_total[df_total['UID'].isin(uids_to_delete)].copy()
                    df_to_delete['刪除時間'] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    df_to_delete = df_to_delete.drop(columns=['move'], errors='ignore')
                    df_remaining = df_total[~df_total['UID'].isin(uids_to_delete)].copy()
                    df_remaining = df_remaining.drop(columns=['move'], errors='ignore')
                    df_remaining = run_cross_matching(df_remaining)
                    df_delete_log = pd.concat([df_delete_log, df_to_delete], ignore_index=True)
                    if save_to_master_xlsm_cloud({"麗嬰國際產品總表": df_remaining, "刪除紀錄": df_delete_log}, MASTER_FILE_ID):
                        st.rerun()
        else:
            st.success("🟢 當前總表中沒有任何重複商品的衝突。")

elif menu == "🔍 麗嬰總表分頁檢視":
    st.header("📋 麗嬰採購產品總表雲端分頁檢視")
    st.metric(label=f"📊 【{selected_sheet}】當前總資料筆數", value=f"{len(df_view)} 筆")
    if search_term:
        search_mask = df_view.astype(str).apply(lambda x: x.str.contains(search_term, case=False, na=False)).any(axis=1)
        st.dataframe(df_view[search_mask], use_container_width=True)
    else:
        st.dataframe(df_view, use_container_width=True)

elif menu == "🧡 蝦皮商品列表優化":
    st.header("🧡 蝦皮商品清單優化系統")
    if uploaded_shopee and btn_shopee:
        df_shopee_raw = pd.read_excel(uploaded_shopee, header=None, engine='openpyxl')
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
            p_idx = cols_list.index("價格")
            cols_list.insert(p_idx, "iSKU")
            df_shopee = df_shopee[cols_list]
        df_valid_isku = df_shopee[df_shopee['iSKU'] != "蝦皮無iSKU"].copy()
        df_isku_keep = df_valid_isku.sort_values(by=['iSKU', '價格', 'original_index']).drop_duplicates(subset=['iSKU'], keep='last')
        df_gtin_check = df_isku_keep.copy()
        df_gtin_check['GTIN_str'] = df_gtin_check['GTIN'].astype(str).str.strip().str.split('.').str[0]
        df_gtin_keep = df_gtin_check[~df_gtin_check['GTIN_str'].isin(["", "00", "0", "nan"])].sort_values(by=['GTIN_str', '價格', 'original_index']).drop_duplicates(subset=['GTIN_str'], keep='last')
        df_final_clean = pd.concat([df_gtin_keep, df_gtin_check[df_gtin_check['GTIN_str'].isin(["", "00", "0", "nan"])]]).sort_values(by='original_index')
        df_final_deleted = df_shopee[~df_shopee.index.isin(df_final_clean.index)]
        for df_tmp in [df_final_clean, df_final_deleted]:
            df_tmp.drop(columns=['original_index', 'GTIN_str'], errors='ignore', inplace=True)
        st.session_state['shopee_clean'] = df_final_clean

    if 'shopee_clean' in st.session_state:
        st.dataframe(st.session_state['shopee_clean'], use_container_width=True)

elif menu == "📊 PowerQuery 三表整合":
    st.header("📊 PowerQuery 雲端商業情報整合中心")
    if btn_pq:
        if not (LOCAL_PROD_ID and LOCAL_SHOPEE_ID):
            st.error("❌ 啟動失敗！請在左側控制台先填入 [商品iSKU清單] 與 [蝦皮商品列表] 的雲端 File ID。")
        else:
            with st.spinner("雲端聯結勾稽中..."):
                try:
                    df_liying = df_total.copy()
                    df_p = pd.read_excel(download_file_from_drive(LOCAL_PROD_ID), sheet_name="商品iSKU清單")
                    df_s = pd.read_excel(download_file_from_drive(LOCAL_SHOPEE_ID), sheet_name="蝦皮商品列表")
                    
                    df_liying['條碼'] = df_liying['條碼'].astype(str).str.strip().str.split('.').str[0]
                    df_p["自定義編碼"] = df_p["自定義編碼"].astype(str).str.strip().str.split('.').str[0]
                    df_s["iSKU"] = df_s["iSKU"].astype(str).str.strip().str.split('.').str[0]
                    
                    df_merge1 = pd.merge(df_p, df_s[["iSKU", "GTIN", "價格"]], left_on="自定義編碼", right_on="iSKU", how="left").rename(columns={"GTIN": "蝦皮GTIN", "價格": "蝦皮售價"})
                    df_merge1["c"] = df_merge1["c"].astype(str).str.strip().str.split('.').str[0]
                    df_final = pd.merge(df_merge1, df_liying[["條碼", "零售價", "含稅"]], left_on="c", right_on="條碼", how="left").rename(columns={"零售價": "麗嬰零售價", "含稅": "麗嬰批發含稅價", "條碼": "麗嬰條碼"})
                    df_final["麗嬰商品"] = df_final["麗嬰條碼"].apply(lambda x: None if pd.isna(x) else "v")
                    for c in ["蝦皮售價", "麗嬰零售價", "麗嬰批發含稅價"]:
                        df_final[c] = pd.to_numeric(df_final[c], errors='coerce')
                    df_final["麗嬰零售八折"] = df_final["麗嬰零售價"] * 0.8
                    df_final["麗嬰八折比蝦皮貴"] = df_final.apply(lambda r: "v" if (pd.notna(r["麗嬰零售八折"]) and pd.notna(r["蝦皮售價"]) and r["麗嬰零售八折"] > r["蝦皮售價"]) else None, axis=1)
                    df_final["麗嬰未稅價"] = df_final["麗嬰批發含稅價"].apply(lambda x: round(x / 1.05) if pd.notna(x) else None)
                    df_final["麗嬰稅款"] = df_final.apply(lambda r: round(r["麗嬰批發含稅價"] - r["麗嬰未稅價"]) if (pd.notna(r["麗嬰批發含稅價"]) and pd.notna(r["麗嬰未稅價"])) else None, axis=1)
                    st.session_state['pq_result'] = df_final.drop(columns=["iSKU"], errors="ignore")
                except Exception as e:
                    st.error(f"❌ 跨表整合失敗: {str(e)}")
                    
    if 'pq_result' in st.session_state:
        st.dataframe(st.session_state['pq_result'], use_container_width=True)

# -------------------------------------------------------------------------
# 🛍️ 模組 5: 🛍️ 採購入庫單模組 (雲端雙向即時存底調閱完全體版)
# -------------------------------------------------------------------------
elif menu == "🛍️ 採購入庫單模組":
    st.header("🛍️ SiteGiant 雲端採購入庫單智慧轉換中心")
    
    inward_tab1, inward_tab2 = st.tabs(["🆕 新建入庫工作區", "📜 雲端紀錄調閱庫"])

    with inward_tab1:
        st.markdown(f"### 📦 當前綁定之單據標頭資訊")
        st.info(f"📅 **收貨日**：`{recv_date}` ｜ 🏬 **指定廠商**：`{vendor_name}` ｜ 📝 **單號**：`{order_no}`")
        
        if bulk_paste_area.strip():
            lines = [line.strip() for line in bulk_paste_area.strip().split('\n') if line.strip()]
            parsed_rows = []
            for line in lines:
                if "," in line: tokens = line.split(',')
                elif "\t" in line: tokens = line.split('\t')
                else: tokens = line.split()
                if tokens:
                    b_code = str(tokens[0]).strip().split('.')[0]
                    q_val = 1 
                    if len(tokens) > 1:
                        try: q_val = int(float(str(tokens[1]).strip()))
                        except: q_val = 1
                    parsed_rows.append({"國際條碼": b_code, "數量": q_val})
            if parsed_rows:
                st.session_state['inward_input_df'] = pd.DataFrame(parsed_rows)
                st.toast("🎯 條碼解析同步成功！")

        input_df = st.data_editor(st.session_state['inward_input_df'], num_rows="dynamic", use_container_width=True, key="inward_grid")
        st.session_state['inward_input_df'] = input_df 
        
        if btn_inward:
            try:
                ref_fh = download_file_from_drive(PRICE_SUMMARY_ID)
                df_ref = pd.read_excel(ref_fh)
                df_ref['c_clean'] = df_ref['c'].astype(str).str.strip().str.split('.').str[0]
                    
                result_rows = []
                for row in input_df.itertuples(index=False):
                    barcode_input = str(row.國際條碼).strip().split('.')[0] if pd.notna(row.國際條碼) else ""
                    if barcode_input in ["", "0", "nan", "None"]: continue
                    qty = int(row.數量) if pd.notna(row.數量) else 0
                    
                    match = df_ref[df_ref['c_clean'] == barcode_input]
                    if not match.empty:
                        match_row = match.iloc[0]
                        result_rows.append({
                            "收貨日": str(recv_date), "國際條碼": barcode_input,
                            "庫存SKU": match_row.get('自定義編碼', '⚠️ 提示：須新增iSKU'),
                            "庫存貨品名稱": match_row.get('品名', match_row.get('名稱', '')),
                            "成本": match_row.get('麗嬰未稅價', '⚠️ 提示：手動輸入'),
                            "稅款": match_row.get('麗嬰稅款', '⚠️ 提示：手動輸入'), "數量": qty,
                            "分類定義": match_row.get('分類定義', ''), "產品關鍵字": match_row.get('產品關鍵字', '')
                        })
                    else:
                        result_rows.append({
                            "收貨日": str(recv_date), "國際條碼": barcode_input,
                            "庫存SKU": "⚠️ 提示：須新增iSKU", "庫存貨品名稱": "請至麗嬰總表補入此新商品", 
                            "成本": "⚠️ 提示：手動輸入", "稅款": "⚠️ 提示：手動輸入", "數量": qty, "分類定義": "", "產品關鍵字": ""
                        })
                        
                if result_rows:
                    st.session_state['inward_result_df'] = pd.DataFrame(result_rows)
                    st.success(f"🚀 SiteGiant 格式勾稽完成！")
            except Exception as e:
                st.error(f"❌ 價格表讀取失敗: {str(e)}")

        if 'inward_result_df' in st.session_state:
            res_df = st.session_state['inward_result_df']
            st.dataframe(res_df, use_container_width=True)
            
            towrite_inward = io.BytesIO()
            with pd.ExcelWriter(towrite_inward, engine='openpyxl') as writer:
                res_df.to_excel(writer, index=False, sheet_name="SiteGiant入庫單")
            towrite_inward.seek(0)
            
            final_filename = f"sitegiant採購入庫單_{vendor_name}_{order_no}.xlsx"
            
            # 🎯【雲端自動存底核心】全自動上傳備份至指定 Google Drive 資料夾
            if st.button("💾 確認此單無誤，一鍵同步備份至雲端歷史夾", type="primary"):
                media_body = MediaFileUpload(towrite_inward, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                file_metadata = {'name': final_filename, 'parents': [HISTORY_FOLDER_ID]}
                drive_service.files().create(body=file_metadata, media_body=media_body, fields='id').execute()
                st.success(f"☁️ 歷史紀錄已完美無縫存底至雲端資料夾中！")

    with inward_tab2:
        st.subheader("🔍 雲端歷史入庫紀錄即時檢索")
        if st.button("🔄 刷新雲端歷史目錄"):
            query = f"'{HISTORY_FOLDER_ID}' in parents and mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'"
            results = drive_service.files().list(q=query, fields="files(id, name)").execute()
            st.session_state['cloud_hist_files'] = results.get('files', [])
            
        if 'cloud_hist_files' in st.session_state and st.session_state['cloud_hist_files']:
            file_options = {f['name']: f['id'] for f in st.session_state['cloud_hist_files']}
            selected_file_name = st.selectbox("🎯 請選擇欲調閱的雲端歷史入庫檔案：", list(file_options.keys()))
            
            if selected_file_name:
                hist_file_id = file_options[selected_file_name]
                df_hist_view = pd.read_excel(download_file_from_drive(hist_file_id))
                st.dataframe(df_hist_view, use_container_width=True)
