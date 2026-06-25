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
# 🌐 1. Google Drive API 雲端連線與 ID 定義 (改用 Streamlit Secrets)
# =========================================================================
# 精準定義您提供的 Google Cloud 雲端硬碟物件 ID 
ID_PROD_FOLDER = "1NtMAYb-SvdH6XMmqB5G07ttB-NuWCqDV"          # 商品列表 資料夾 ID
ID_PRICE_SUMMARY_FOLDER = "1ZM4MscX0UO6rUHjKv-mN5fKDwxg53maZ" # 價格統整表 資料夾 ID
ID_SHOPEE_FOLDER = "17eiGnXyU4KwNS6IR5bubBPti46SKXMH0"        # 蝦皮商品清單 資料夾 ID
ID_HISTORY_INWARD_FOLDER = "1ZQ7x4BdRc6BJlURxQ61JqDKrKF7h_vSH"# 歷史入庫單 資料夾 ID
ID_BASE_FOLDER = "1HjMt8z8DXlqGhSqe50_hDR3f4LpVLK_w"          # 麗嬰採購統整 資料夾 ID

@st.cache_resource
def init_drive_service():
    """讀取部署後設定在 Streamlit Secrets 的金鑰字典並建立連線"""
    try:
        google_secrets = st.secrets["textkey"]
        credentials = service_account.Credentials.from_service_account_info(
            google_secrets,
            scopes=["https://www.googleapis.com/auth/drive"]
        )
        return build('drive', 'v3', credentials=credentials)
    except Exception as e:
        st.error(f"❌ 無法從 Streamlit Secrets 中讀取 `textkey` 憑證。錯誤訊息: {str(e)}")
        st.info("💡 請確認您的 Secrets 設定格式是否正確（必須包含 textkey 區塊）。")
        st.stop()

# 呼叫初始化函式取得服務個體
service = init_drive_service()

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
    except Exception:
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
    """將 Google API 返回的 ISO 時間格式轉換為易讀格式 (台灣時間)"""
    if not time_str:
        return "❌ 雲端檔案尚未建立/不存在"
    try:
        dt = datetime.datetime.strptime(time_str.split('.')[0], "%Y-%m-%dT%H:%M:%S")
        dt = dt + datetime.timedelta(hours=8)
        return dt.strftime("%Y-%m-%d %H:%M:%S (台灣時間)")
    except:
        return time_str

# -------------------------------------------------------------------------
# 動態偵測各核心主表的雲端 ID ＆ 最後修改時間
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
            df_delete_log = pd.read_excel(xls, "刪除紀錄") if "刪除紀錄" in xls.sheet_names else pd.DataFrame(columns=["UID", "名稱", "條碼", "零售價", "備註", "匯入档名", "刪除時間"])
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
