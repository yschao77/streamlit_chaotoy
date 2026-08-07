import streamlit as st
import pandas as pd
import openpyxl
import io
import datetime
import json
import os

# 🌟 從我們剛建立的 utils 中匯入所有工具函式
from utils import (
    apply_openpyxl_patch, HAS_CALAMINE, get_cached_gdrive_id, 
    list_gdrive_files, download_gdrive_file_to_bytes, 
    get_cached_gdrive_file_bytes, upload_or_update_gdrive_file, 
    format_gdrive_time, calculate_md5, clean_barcode, process_smart_headers
)

from views import sitegiant_page, integration_page

# 執行修補程式
apply_openpyxl_patch()

# 設定網頁標題與寬度
st.set_page_config(page_title="麗嬰與蝦皮商務數據情報中心", page_icon="📊", layout="wide")

# =========================================================================
# 🌐 雲端資料夾 ID 定義 (這些是常數，留在主檔沒問題)
# =========================================================================
ID_PROD_FOLDER = "1NtMAYb-SvdH6XMmqB5G07ttB-NuWCqDV"          
ID_PRICE_SUMMARY_FOLDER = "1ZM4MscX0UO6rUHjKv-mN5fKDwxg53maZ" 
ID_SHOPEE_FOLDER = "17eiGnXyU4KwNS6IR5bubBPti46SKXMH0"        
ID_HISTORY_INWARD_FOLDER = "1ZQ7x4BdRc6BJlURxQ61JqDKrKF7h_vSH"
ID_BASE_FOLDER = "1HjMt8z8DXlqGhSqe50_hDR3f4LpVLK_w"          

# 取得主表 ID 狀態
ID_MASTER_FILE, TIME_MASTER, NAME_MASTER = get_cached_gdrive_id(ID_BASE_FOLDER, "麗嬰採購產品總表")
ID_LOCAL_PROD, TIME_PROD, NAME_PROD = get_cached_gdrive_id(ID_PROD_FOLDER, "商品列表")
ID_SHOPEE_MASTER, TIME_SHOPEE, NAME_SHOPEE = get_cached_gdrive_id(ID_SHOPEE_FOLDER, "蝦皮賣場商品列表")
ID_PRICE_SUMMARY, TIME_SUMMARY, NAME_SUMMARY = get_cached_gdrive_id(ID_PRICE_SUMMARY_FOLDER, "商品蝦皮麗嬰價格統整表")

# =========================================================================
# ⚙️ 3. 核心資料庫讀寫雲端常式
# =========================================================================
def save_to_master_xlsm(sheets_dict):
    if not ID_MASTER_FILE:
        st.error(f"❌ 雲端找不到核心總表檔案 【麗嬰採購產品總表.xlsm】")
        return False
    try:
        master_bytes = download_gdrive_file_to_bytes(ID_MASTER_FILE)
        wb = openpyxl.load_workbook(master_bytes, keep_vba=True)
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
                    elif barcode_col_idx and (col_idx + 1) == barcode_col_idx: cleaned_row.append(str(x).strip().split('.')[0])
                    else: cleaned_row.append(x)
                ws.append(cleaned_row)
                if barcode_col_idx: ws.cell(row=row_idx, column=barcode_col_idx).number_format = '@'
                    
        out_buf = io.BytesIO()
        wb.save(out_buf)
        upload_or_update_gdrive_file(ID_BASE_FOLDER, NAME_MASTER or "麗嬰採購產品總表.xlsm", out_buf.getvalue(), existing_file_id=ID_MASTER_FILE)
        return True
    except Exception as e:
        st.error(f"❌ 寫入雲端資料庫發生錯誤: {str(e)}")
        return False

def save_to_shopee_master_xlsm(sheets_dict):
    global ID_SHOPEE_MASTER
    try:
        if ID_SHOPEE_MASTER:
            shopee_bytes = download_gdrive_file_to_bytes(ID_SHOPEE_MASTER)
            wb = openpyxl.load_workbook(shopee_bytes, keep_vba=True)
        else:
            wb = openpyxl.Workbook()
            
        for sheet_name, df in sheets_dict.items():
            if sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                ws.delete_rows(1, ws.max_row + 1)
            else:
                ws = wb.create_sheet(sheet_name)
            ws.append(list(df.columns))
            for row in df.itertuples(index=False):
                cleaned_row = [" " if pd.isna(x) else x for x in row]
                ws.append(cleaned_row)
                
        out_buf = io.BytesIO()
        wb.save(out_buf)
        ID_SHOPEE_MASTER = upload_or_update_gdrive_file(
            ID_SHOPEE_FOLDER, 
            NAME_SHOPEE or "蝦皮賣場商品列表.xlsm", 
            out_buf.getvalue(), 
            existing_file_id=ID_SHOPEE_MASTER
        )
        return True
    except Exception as e:
        st.error(f"❌ 寫入雲端蝦皮資料庫發生錯誤: {str(e)}")
        return False

def run_cross_matching(df):
    if df.empty: return df
    df['備註'] = df['備註'].astype(str).apply(lambda x: "" if x == "nan" or x == "None" else x)
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
                existing_remark = str(check_row.get('備註', '')).strip()
                if existing_remark and not any(k in existing_remark for k in ["條碼重複", "名稱不同", "零售價不同"]):
                    continue
                    
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
# ⚙️ 4. 全域雲端資料庫載入與初始化安全驗證 (優化：全面快取與延遲載入)
# =========================================================================
@st.cache_data(ttl=600)
def load_master_data(file_id):
    """延遲載入並快取麗嬰主表資料"""
    if not file_id: return None, None, None, None, None, 3473
    try:
        master_bytes = download_gdrive_file_to_bytes(file_id)
        # 優化：此處若不寫入，僅供查詢，建議預設使用 openpyxl 讀取保留相容性，但查詢頁面可改用 calamine
        with pd.ExcelFile(master_bytes) as xls:
            df_total = pd.read_excel(xls, "麗嬰國際產品總表")
            df_history = pd.read_excel(xls, "已處理採購單")
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
            if len(df_history.columns) >= 3:
                df_history.columns = standard_cols + list(df_history.columns[3:])
            else:
                df_history = pd.DataFrame(columns=standard_cols)
        else:
            df_history = pd.DataFrame(columns=["檔案名稱", "md5", "匯入時間"])
            
        current_max_uid = int(df_meta.iloc[0, 0]) if not df_meta.empty else 3473
        return df_total, df_history, df_delete_log, df_meta, all_sheets, current_max_uid
    except Exception as e:
        st.error(f"🔴 讀取雲端主資料庫失敗。錯誤: {str(e)}")
        return None, None, None, None, None, 3473

@st.cache_data(ttl=600)
def load_shopee_data(file_id):
    """延遲載入並快取蝦皮主表資料"""
    if not file_id: return pd.DataFrame(columns=["檔案名稱", "md5", "匯入時間"]), pd.DataFrame()
    try:
        shopee_bytes = download_gdrive_file_to_bytes(file_id)
        with pd.ExcelFile(shopee_bytes) as shopee_xls:
            df_hist = pd.read_excel(shopee_xls, "匯入檔案") if "匯入檔案" in shopee_xls.sheet_names else pd.DataFrame(columns=["檔案名稱", "md5", "匯入時間"])
            df_list = pd.read_excel(shopee_xls, "蝦皮商品列表") if "蝦皮商品列表" in shopee_xls.sheet_names else pd.DataFrame()
        return df_hist, df_list
    except Exception:
        return pd.DataFrame(columns=["檔案名稱", "md5", "匯入時間"]), pd.DataFrame()

if 'inward_input_df' not in st.session_state:
    st.session_state['inward_input_df'] = pd.DataFrame([{"國際條碼": "", "數量": 1}])

# --- 條碼格式終極清洗函數 ---
def clean_barcode(val):
    """強制清洗條碼：去除空格、網頁空白字元、還原 E+ 科學記號、去除浮點數後綴"""
    if pd.isna(val): 
        return ""
    s = str(val).strip().replace(" ", "").replace("\xa0", "")
    if s.lower() == "nan" or s == "": 
        return ""
    if s.endswith(".0"): 
        s = s[:-2]
    # 還原 Pandas 誤讀的科學記號 (如 4.71E+12)
    if "e+" in s.lower() or "e" in s.lower():
        try:
            s = f"{float(s):.0f}"
        except:
            pass
    return s

# --- 智慧複合表頭處理函數 ---
def process_smart_headers(df_raw, header_row_idx):
    """完美重現 VBA 邏輯：當第5列有特定字眼時，與第6列合併為新標頭"""
    if header_row_idx >= 5 and len(df_raw) > 5:
        row_5 = df_raw.iloc[4].fillna("").astype(str).str.strip()
        row_6 = df_raw.iloc[5].fillna("").astype(str).str.strip()
        new_cols = []
        for idx in range(len(df_raw.columns)):
            c5 = row_5.iloc[idx] if idx < len(row_5) else ""
            c6 = row_6.iloc[idx] if idx < len(row_6) else ""
            
            # 若第5列包含採購關鍵字，且與第6列不同，則動態結合 (例如: 內盒 + 12 = 內盒12)
            if any(k in c5 for k in ["訂購", "CTN", "內盒", "數量", "單價"]) and c5 != c6 and c6 != "":
                new_cols.append(f"{c5}{c6}")
            elif c6 != "":
                new_cols.append(c6)
            else:
                new_cols.append(c5)
        return new_cols
    else:
        return [str(col).strip() for col in df_raw.iloc[header_row_idx].fillna("")]

# ==========================================
# 🧭 5. 側邊欄：導覽控制台
# ==========================================
st.sidebar.markdown("## 🏢 進銷存中央管理系統")
st.sidebar.write("---")

main_module = st.sidebar.selectbox(
    "🎯 請選擇核心管理模組：",
    ["🏪 Sitegiant 電商整合管理", "📦 商品蝦皮麗嬰統整管理"]
)
st.sidebar.write("") 

# -------------------------------------------------------------------------
# 打包所有的雲端 ID 與參數給子頁面使用
# -------------------------------------------------------------------------
gdrive_cfg = {
    "ID_BASE_FOLDER": ID_BASE_FOLDER,
    "ID_MASTER_FILE": ID_MASTER_FILE,
    "TIME_MASTER": TIME_MASTER,
    "NAME_MASTER": NAME_MASTER,
    
    "ID_PROD_FOLDER": ID_PROD_FOLDER,
    "ID_LOCAL_PROD": ID_LOCAL_PROD,
    "TIME_PROD": TIME_PROD,
    "NAME_PROD": NAME_PROD,
    
    "ID_SHOPEE_FOLDER": ID_SHOPEE_FOLDER,
    "ID_SHOPEE_MASTER": ID_SHOPEE_MASTER,
    "TIME_SHOPEE": TIME_SHOPEE,
    "NAME_SHOPEE": NAME_SHOPEE,
    
    "ID_PRICE_SUMMARY_FOLDER": ID_PRICE_SUMMARY_FOLDER,
    "ID_PRICE_SUMMARY": ID_PRICE_SUMMARY,
    "TIME_SUMMARY": TIME_SUMMARY,
    "NAME_SUMMARY": NAME_SUMMARY
}

# ==========================================
# 🧭 5. 側邊欄：導覽控制台
# ==========================================
# (這裡保留你剛才修好的側邊欄與頁面路由)
if "商品蝦皮麗嬰統整管理" in main_module:
    st.sidebar.markdown("### 🛠️ 整合合併轉換功能")
    sub_page = st.sidebar.radio(
        "請選擇執行項目：",
        ["🧠 PowerQuery 執行三表整合", "⚖️ 麗嬰商品表合併和與審核", "📈 蝦皮商品清單轉換", "📊 PowerQuery 三表整合歷史紀錄", "🔍 商品清單紀錄查詢", "🔍 麗嬰商品總表數據查詢"],
        index=3
    )
    # 👇 呼叫我們剛剛新建的整合頁面模組
    integration_page.render(sub_page, gdrive_cfg)

else:
    st.sidebar.markdown("### 🌐 Sitegiant 電商整合管理")
    sub_page = st.sidebar.radio(
        "請選擇執行項目：",
<<<<<<< HEAD
        ["🔀 Sitegiant 採購入庫單格式轉換", "📜 Sitegiant 歷史入庫單紀錄", "📦 Sitegiant 批量新增UPC", "📋 採購單待處理"],
=======
        ["🔀 Sitegiant 採購入庫單格式轉換", "📜 Sitegiant 歷史入庫單紀錄", "📦 Sitegiant 批量新增UPC", "📋 採購單待處理"], 
>>>>>>> a4b33deec2535672f61b8e39fb784203a8c4e5a0
        index=0
    )
    sitegiant_page.render(
        sub_page=sub_page, 
        ID_PRICE_SUMMARY=ID_PRICE_SUMMARY, 
        ID_HISTORY_INWARD_FOLDER=ID_HISTORY_INWARD_FOLDER, 
        ID_SHOPEE_MASTER=ID_SHOPEE_MASTER
    )
