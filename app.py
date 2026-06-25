import streamlit as st
import pandas as pd
import openpyxl
import hashlib
import datetime
import io
import os
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# 設定網頁標題與寬度
st.set_page_config(page_title="麗嬰與蝦皮商務數據情報中心", page_icon="📊", layout="wide")

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
# 🌐 1. Google Drive API 雲端連線與 ID 定義
# =========================================================================
SCOPES = ['https://www.googleapis.com/auth/drive']

# 精準定義您提供的 Google Cloud 雲端硬碟物件 ID 
ID_PROD_FOLDER = "1NtMAYb-SvdH6XMmqB5G07ttB-NuWCqDV"          # 商品列表 資料夾 ID
ID_PRICE_SUMMARY_FOLDER = "1ZM4MscX0UO6rUHjKv-mN5fKDwxg53maZ" # 價格統整表 資料夾 ID
ID_SHOPEE_FOLDER = "17eiGnXyU4KwNS6IR5bubBPti46SKXMH0"        # 蝦皮商品清單 資料夾 ID
ID_HISTORY_INWARD_FOLDER = "1ZQ7x4BdRc6BJlURxQ61JqDKrKF7h_vSH"# 歷史入庫單 資料夾 ID
ID_BASE_FOLDER = "1HjMt8z8DXlqGhSqe50_hDR3f4LpVLK_w"          # 麗嬰採購統整 資料夾 ID

@st.cache_resource
def get_gdrive_service():
    """從環境變數讀取憑證並建立 Google Drive 服務連線"""
    gdrive_json_str = os.environ.get("GDRIVE_CREDENTIALS_JSON")
    if gdrive_json_str:
        creds_dict = json.loads(gdrive_json_str)
        creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    else:
        # 本地測試備用方案
        SERVICE_ACCOUNT_FILE = 'gdrive_credentials.json'
        if os.path.exists(SERVICE_ACCOUNT_FILE):
            creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        else:
            st.error("❌ 找不到 Google Drive 雲端憑證 (GDRIVE_CREDENTIALS_JSON)！")
            st.stop()
    return build('drive', 'v3', credentials=creds)

service = get_gdrive_service()

# =========================================================================
# 🔍 2. 雲端檔案搜尋與讀寫核心工具函式
# =========================================================================
def find_gdrive_file(folder_id, file_name_keyword):
    """在指定的 Google Drive 資料夾中，根據名稱搜尋檔案並返回其 ID 與修改時間"""
    try:
        query = f"'{folder_id}' in parents and name contains '{file_name_keyword}' and trashed = false"
        results = service.files().list(q=query, fields="files(id, name, modifiedTime)", pageSize=1).execute()
        files = results.get('files', [])
        if files:
            return files[0]['id'], files[0]['modifiedTime'], files[0]['name']
    except Exception as e:
        pass
    return None, None, None

def list_gdrive_files(folder_id):
    """列出雲端資料夾內的所有 Excel 檔案 (用於調閱歷史紀錄)"""
    try:
        query = f"'{folder_id}' in parents and (mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or mimeType = 'application/vnd.ms-excel.sheet.macroEnabled.12') and trashed = false"
        results = service.files().list(q=query, fields="files(id, name, modifiedTime)").execute()
        return results.get('files', [])
    except Exception as e:
        st.error(f"掃描雲端資料夾失敗: {str(e)}")
        return []

def download_gdrive_file_to_bytes(file_id):
    """將雲端檔案下載至記憶體 (BytesIO) 中，不落地存成實體檔案"""
    request = service.files().get_media(fileId=file_id)
    file_stream = io.BytesIO()
    downloader = MediaIoBaseDownload(file_stream, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
    file_stream.seek(0)
    return file_stream

def upload_or_update_gdrive_file(folder_id, file_name, file_bytes, existing_file_id=None):
    """上傳新檔案到 Google Drive 或更新現有的檔案"""
    media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
    if existing_file_id:
        # 覆寫更新現有檔案
        service.files().update(fileId=existing_file_id, media_body=media).execute()
        return existing_file_id
    else:
        # 新增檔案至指定資料夾
        file_metadata = {'name': file_name, 'parents': [folder_id]}
        file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        return file.get('id')

def format_gdrive_time(time_str):
    """將 Google API 返回的 ISO 時間格式轉換為易讀格式"""
    if not time_str:
        return "❌ 雲端檔案尚未建立/不存在"
    try:
        dt = datetime.datetime.strptime(time_str.split('.')[0], "%Y-%m-%dT%H:%M:%S")
        # 加上 8 小時轉為台灣時間
        dt = dt + datetime.timedelta(hours=8)
        return dt.strftime("%Y-%m-%d %H:%M:%S (台灣時間)")
    except:
        return time_str

# -------------------------------------------------------------------------
# 動態偵測三個核心主表的雲端 ID ＆ 最後修改時間
# -------------------------------------------------------------------------
ID_MASTER_FILE, TIME_MASTER, NAME_MASTER = find_gdrive_file(ID_BASE_FOLDER, "麗嬰採購產品總表")
ID_LOCAL_PROD, TIME_PROD, NAME_PROD = find_gdrive_file(ID_PROD_FOLDER, "商品列表")
ID_SHOPEE_MASTER, TIME_SHOPEE, NAME_SHOPEE = find_gdrive_file(ID_SHOPEE_FOLDER, "蝦皮賣場商品列表")
ID_PRICE_SUMMARY, TIME_SUMMARY, NAME_SUMMARY = find_gdrive_file(ID_PRICE_SUMMARY_FOLDER, "商品蝦皮麗嬰價格統整表")

# =========================================================================
# ⚙️ 3. 雲端資料庫全域資料預先連線載入
# =========================================================================
if ID_MASTER_FILE:
    try:
        master_bytes = download_gdrive_file_to_bytes(ID_MASTER_FILE)
        with pd.ExcelFile(master_bytes) as xls:
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
            if len(df_history.columns) >= 3:
                df_history.columns = standard_cols + list(df_history.columns[3:])
            else:
                df_history = pd.DataFrame(columns=standard_cols)
        else:
            df_history = pd.DataFrame(columns=["檔案名稱", "md5", "匯入時間"])
        current_max_uid = int(df_meta.iloc[0, 0]) if not df_meta.empty else 3473
    except Exception as e:
        st.error(f"🔴 讀取雲端主資料庫失敗，請確認資料夾內是否有『麗嬰採購產品總表』。錯誤: {str(e)}")
        st.stop()
else:
    st.error(f"❌ 於雲端資料庫路徑下找不到核心主檔案 【麗嬰採購產品總表.xlsm】，請先上傳至 Google Drive。")
    st.stop()

# 載入雲端蝦皮資料庫 Master
try:
    if ID_SHOPEE_MASTER:
        shopee_bytes = download_gdrive_file_to_bytes(ID_SHOPEE_MASTER)
        with pd.ExcelFile(shopee_bytes) as shopee_xls:
            df_shopee_history = pd.read_excel(shopee_xls, "匯入檔案") if "匯入檔案" in shopee_xls.sheet_names else pd.DataFrame(columns=["檔案名稱", "md5", "匯入時間"])
            df_shopee_current_list = pd.read_excel(shopee_xls, "蝦皮商品列表") if "蝦皮商品列表" in shopee_xls.sheet_names else pd.DataFrame()
    else:
        df_shopee_history = pd.DataFrame(columns=["檔案名稱", "md5", "匯入時間"])
        df_shopee_current_list = pd.DataFrame()
except Exception:
    df_shopee_history = pd.DataFrame(columns=["檔案名稱", "md5", "匯入時間"])
    df_shopee_current_list = pd.DataFrame()

if 'inward_input_df' not in st.session_state:
    st.session_state['inward_input_df'] = pd.DataFrame([{"國際條碼": "", "數量": 1}])

# ==========================================
# ⚙️ 4. 寫入雲端檔案的核心實戰常式 (取代原本本地 wb.save)
# ==========================================
def save_to_master_xlsm(sheets_dict):
    try:
        # 先下載目前的雲端範本 (保留巨集與結構)
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
                    if pd.isna(x):
                        cleaned_row.append("")
                    elif barcode_col_idx and (col_idx + 1) == barcode_col_idx:
                        cleaned_row.append(str(x).strip().split('.')[0])
                    else:
                        cleaned_row.append(x)
                ws.append(cleaned_row)
                if barcode_col_idx:
                    ws.cell(row=row_idx, column=barcode_col_idx).number_format = '@'
                    
        # 將修改後的 Workbook 存入 memory buffer 中並回傳至 Google Drive 覆寫
        out_buf = io.BytesIO()
        wb.save(out_buf)
        upload_or_update_gdrive_file(ID_BASE_FOLDER, NAME_MASTER, out_buf.getvalue(), existing_file_id=ID_MASTER_FILE)
        return True
    except Exception as e:
        st.error(f"❌ 寫入雲端資料庫發生錯誤: {str(e)}")
        return False

def save_to_shopee_master_xlsm(sheets_dict):
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
        # 更新或創建雲端蝦皮總表
        global ID_SHOPEE_MASTER
        ID_SHOPEE_MASTER = upload_or_update_gdrive_file(ID_SHOPEE_FOLDER, NAME_SHOPEE or "蝦皮賣場商品列表.xlsm", out_buf.getvalue(), existing_file_id=ID_SHOPEE_MASTER)
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
                try:
                    price_val = float(check_row.get('零售價', 0)) if pd.notna(check_row.get('零售價', 0)) else 0
                except:
                    price_val = 0
                for comp_idx in indices:
                    if check_idx == comp_idx: continue
                    comp_row = records[comp_idx]
                    comp_name_str = str(comp_row.get('名稱', '')).strip()
                    try:
                        comp_price_val = float(comp_row.get('零售價', 0)) if pd.notna(comp_row.get('零售價', 0)) else 0
                    except:
                        comp_price_val = 0
                    uid_str = str(comp_row.get('UID', '未知')).strip()
                    if name_str != comp_name_str:
                        records[check_idx]['備註'] = f"與 UID: {uid_str} 條碼重複, 名稱不同"
                        break
                    elif price_val != comp_price_val:
                        records[check_idx]['備註'] = f"與 UID: {uid_str} 條碼重複, 名稱相同, 零售價不同"
                        break
    return pd.DataFrame(records)

# ==========================================
# 🧭 側邊欄：視覺化模組導覽
# ==========================================
st.sidebar.markdown("## 🏢 雲端進銷存中央管理系統")
st.sidebar.write("---")

main_module = st.sidebar.selectbox(
    "🎯 請選擇核心管理模組：",
    ["📦 商品蝦皮麗嬰統整管理", "🏪 sitegiant 電商整合管理"]
)
st.sidebar.write("") 

if "商品蝦皮麗嬰統整管理" in main_module:
    st.sidebar.markdown("### 🛠️ 整合合併轉換功能")
    sub_page = st.sidebar.radio(
        "請選擇執行項目：",
        [
            "🧠 PowerQuery 執行三表整合",
            "⚖️ 麗嬰商品表合併和與審核",
            "📈 蝦皮商品清單轉換",
            "📊 PowerQuery 三表整合歷史紀錄",
            "🔍 商品清單紀錄查詢",
            "🔍 麗嬰商品總表數據查詢",
         ],
        index=3
    )
else:
    st.sidebar.markdown("### 🌐 sitegiant 電商整合管理")
    sub_page = st.sidebar.radio(
        "請選擇執行項目：",
        ["🔀 sitegiant 採購入庫單格式轉換", "📜 sitegiant 歷史入庫單紀錄"],
        index=0
    )

# ==========================================
# 🖥️ 主畫面內容呈現與邏輯串接
# ==========================================
st.title(f"{sub_page}")
st.info(f"雲端導覽路徑： {main_module} ➔ {sub_page}")
st.write("---")

# -------------------------------------------------------------------------
# 子功能 1：📊 三表整合歷史 (雲端版本)
# -------------------------------------------------------------------------
if sub_page == "📊 PowerQuery 三表整合歷史紀錄":
    st.subheader("🔄 三表整合歷史紀錄追蹤")
    st.caption("系統會即時掃描 Google Drive `價格統整表/` 資料夾，方便您快速查看與下載過往比對清單。")
    
    hist_pq_files = list_gdrive_files(ID_PRICE_SUMMARY_FOLDER)
    
    if not hist_pq_files:
        st.warning(f"💡 提示：目前雲端資料夾內尚無任何歷史檔案，請至『🧠 PowerQuery 三表整合』執行新建轉換。")
    else:
        # 建立選單
        file_options = {f['name']: f['id'] for f in hist_pq_files}
        selected_pq_file = st.selectbox("🎯 請選擇欲調閱的雲端歷史整合報告：", list(file_options.keys()))
        
        if selected_pq_file:
            try:
                target_id = file_options[selected_pq_file]
                file_bytes = download_gdrive_file_to_bytes(target_id)
                df_pq_view = pd.read_excel(file_bytes)
                
                st.markdown(f"📄 **目前調閱雲端檔案**：`{selected_pq_file}` ｜ 📊 **資料總項數**：`{len(df_pq_view)} 筆`")
                st.dataframe(df_pq_view, use_container_width=True)
                
                st.download_button(
                    label="🔄 重新下載此歷史整合表 (.xlsx)",
                    data=file_bytes.getvalue(),
                    file_name=selected_pq_file,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"❌ 讀取雲端備份檔案失敗: {str(e)}")

# -------------------------------------------------------------------------
# 子功能 2：🧠 PowerQuery 三表整合 (雲端版本)
# -------------------------------------------------------------------------
elif sub_page == "🧠 PowerQuery 執行三表整合":
    st.subheader("🔍 三表數據追蹤")
    st.info(f"☁️ 目前正在運作於： **Google Cloud 雲端直連模式**")    
    
    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("📦 商品列表 (商品iSKU清單)", "已對接 (Cloud)" if ID_LOCAL_PROD else "❌ 雲端未偵測到")
        st.caption(f"📅 最後修改時間: \n`{format_gdrive_time(TIME_PROD)}`")
    with c2:
        st.metric("🧡 蝦皮資料庫主表 (蝦皮商品列表)", "已對接 (Cloud)" if ID_SHOPEE_MASTER else "❌ 雲端未偵測到")
        st.caption(f"📅 最後修改時間: \n`{format_gdrive_time(TIME_SHOPEE)}`")
    with c3:
        st.metric("🧸 麗嬰產品總表", "已對接 (Cloud)" if ID_MASTER_FILE else "❌ 雲端未偵測到")
        st.caption(f"📅 最後修改時間: \n`{format_gdrive_time(TIME_MASTER)}`")

    st.write("---")

    if st.button("🛠️ 啟動三表整合與財務指標計算", type="primary", use_container_width=True):
        if not (ID_LOCAL_PROD and ID_SHOPEE_MASTER and ID_MASTER_FILE):
            st.error("❌ 無法啟動整合！請檢查上方儀表板確認三個雲端主表是否均已上傳至 Google Drive。")
        else:
            with st.spinner("正在直接由雲端載入數據流並進行大數據跨表計算..."):
                try:
                    df_liying = pd.read_excel(download_gdrive_file_to_bytes(ID_MASTER_FILE), sheet_name="麗嬰國際產品總表")
                    df_p = pd.read_excel(download_gdrive_file_to_bytes(ID_LOCAL_PROD), sheet_name="商品iSKU清單")
                    df_s = pd.read_excel(download_gdrive_file_to_bytes(ID_SHOPEE_MASTER), sheet_name="蝦皮商品列表")
                    
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
                    df_final["麗嬰稅款"] = df_final.apply(lambda r: round(r["麗嬰批發含稅價"] - r["麗嬰未稅價"]) if (pd.notna(r["麗批發含稅價"]) and pd.notna(r["麗嬰未稅價"])) else None, axis=1)
                    
                    # 輸出報表至雲端備份
                    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    backup_filename = f"商品蝦皮麗嬰價格統整表_{timestamp}.xlsx"
                    
                    out_buf = io.BytesIO()
                    df_final.to_excel(out_buf, index=False)
                    
                    # 1. 儲存至歷史紀錄資料夾
                    upload_or_update_gdrive_file(ID_PRICE_SUMMARY_FOLDER, backup_filename, out_buf.getvalue())
                    # 2. 覆寫更新最新版的主表 (如果原本存在就傳 ID，不存在則傳 None 新建)
                    upload_or_update_gdrive_file(ID_PRICE_SUMMARY_FOLDER, "商品蝦皮麗嬰價格統整表.xlsx", out_buf.getvalue(), existing_file_id=ID_PRICE_SUMMARY)
                    
                    st.session_state['pq_result'] = df_final.drop(columns=["iSKU"], errors="ignore")
                    st.success(f"🎉 雲端三表整合與財務指標計算完成！已成功備份至雲端：{backup_filename}")
                except Exception as e:
                    st.error(f"❌ 錯誤: {str(e)}")

    if 'pq_result' in st.session_state:
        st.subheader("📋 整合聯結情報報表輸出預覽")
        st.dataframe(st.session_state['pq_result'], use_container_width=True)
        
        towrite_pq = io.BytesIO()
        with pd.ExcelWriter(towrite_pq, engine='openpyxl') as writer:
            st.session_state['pq_result'].to_excel(writer, index=False, sheet_name="PowerQuery三表整合")
        st.download_button(
            label="📥 下載本次整合比對表 (.xlsx)",
            data=towrite_pq.getvalue(),
            file_name=f"三表整合比對結果_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# -------------------------------------------------------------------------
# 子功能 3：🔍 麗嬰商品總表數據查詢
# -------------------------------------------------------------------------
elif sub_page == "🔍 麗嬰商品總表數據查詢":
    st.subheader("📋 麗嬰採購產品總表資料庫分頁動態檢視")
    st.info(f"☁️ 目前雲端主資料庫修改時間：{format_gdrive_time(TIME_MASTER)}")
    
    view_sheets = [s for s in all_sheets if s != "麗嬰產品新採購單"]
    selected_sheet = st.selectbox("請選擇數據分頁：", view_sheets)
    search_term = st.text_input("🔍 快速搜尋關鍵字 (支援條碼、品名、貨號模糊比對)：", placeholder="輸入搜尋內容...")
    
    try:
        df_view = pd.read_excel(download_gdrive_file_to_bytes(ID_MASTER_FILE), selected_sheet)
        if selected_sheet == "麗嬰國際產品總表" and "條碼" in df_view.columns:
            df_view['條碼'] = df_view['條碼'].astype(str).str.strip().str.split('.').str[0]
        st.metric(label=f"📊 【{selected_sheet}】當前總資料筆數", value=f"{len(df_view)} 筆")
        if search_term:
            search_mask = df_view.astype(str).apply(lambda x: x.str.contains(search_term, case=False, na=False)).any(axis=1)
            st.dataframe(df_view[search_mask], use_container_width=True)
        else:
            st.dataframe(df_view, use_container_width=True)
    except Exception as e:
        st.error(f"❌ 讀取分頁數據失敗: {str(e)}")

# -------------------------------------------------------------------------
# 子功能 5：🔍 商品清單紀錄查詢 (雲端版本)
# -------------------------------------------------------------------------
elif sub_page == "🔍 商品清單紀錄查詢":
    st.subheader("📊 歷史商品清單紀錄查詢")
    st.caption("系統會自動追蹤雲端 `商品列表/` 資料夾下的文件，為您呈現目前最新商品清單iSKU及分類。")

    hist_files = list_gdrive_files(ID_PROD_FOLDER)
    
    if not hist_files:
        st.warning(f"💡 目前雲端 `商品列表/` 內無歷史單據。")
    else:
        file_options = {f['name']: f['id'] for f in hist_files}
        selected_hist_file = st.selectbox("🎯 商品清單紀錄:", list(file_options.keys()))
        
        if selected_hist_file:
            try:
                target_id = file_options[selected_hist_file]
                file_bytes = download_gdrive_file_to_bytes(target_id)
                df_hist_view = pd.read_excel(file_bytes)
                st.markdown(f"📄 **當前檔案**：`{selected_hist_file}` ｜ 📊 **iSKU品項數**：`{len(df_hist_view)} 筆`")
                st.dataframe(df_hist_view, use_container_width=True)
            except Exception as e:
                st.error(f"❌ 讀取失敗: {str(e)}")

# (提示：後續的合併審核、sitegiant轉換等子功能，在原本寫入本地端檔案的位置「wb.save()」之處，也都可以直接套用上面定義好的「save_to_master_xlsm」或「upload_or_update_gdrive_file」工具函式來完成雲端同步。)
