import streamlit as st
import pandas as pd
import openpyxl
import hashlib
import datetime
import io
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# =========================================================================
# 🛠️ 🔴 全域攔截並修補 openpyxl 描述器核心驗證 Bug (Monkey Patch)
# =========================================================================
def apply_openpyxl_patch():
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
# 🛠️ check python-calamine
# =========================================================================
def check_calamine():
    try:
        import calamine
        return True
    except ImportError:
        return False

HAS_CALAMINE = check_calamine()

# =========================================================================
# 🌐 1. Google Drive 雲端連線初始化
# =========================================================================
@st.cache_resource
def init_drive_service():
    """讀取部署後設定在 Streamlit Secrets 的金鑰字典並建立雲端連線"""
    try:
        google_secrets = st.secrets["textkey"]
        credentials = service_account.Credentials.from_service_account_info(
            google_secrets,
            scopes=["https://www.googleapis.com/auth/drive"]
        )
        return build('drive', 'v3', credentials=credentials)
    except Exception as e:
        st.error(f"❌ 無法從 Streamlit Secrets 中讀取 `textkey` 憑證。錯誤訊息: {str(e)}")
        st.info("💡 請確認您的 Secrets 設定格式是否正確。")
        st.stop()

# 建立全域 service 供下方工具使用
service = init_drive_service()

# =========================================================================
# 🔍 2. 雲端核心實戰工具與搜尋常式
# =========================================================================
@st.cache_data(ttl=3600)
def get_cached_gdrive_id(folder_id, file_name_keyword):
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
    try:
        query = f"'{folder_id}' in parents and (mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or mimeType = 'application/vnd.ms-excel.sheet.macroEnabled.12') and trashed = false"
        results = service.files().list(q=query, fields="files(id, name, modifiedTime)").execute()
        files = results.get('files', [])
        files.sort(key=lambda x: x['name'], reverse=True)
        return files
    except Exception as e:
        st.error(f"掃描雲端資料夾失敗: {str(e)}")
        return []

def download_gdrive_file_to_bytes(file_id):
    request = service.files().get_media(fileId=file_id)
    file_stream = io.BytesIO()
    downloader = MediaIoBaseDownload(file_stream, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
    file_stream.seek(0)
    return file_stream

@st.cache_data(ttl=600, show_spinner="☁️ 正在從雲端載入檔案...")
def get_cached_gdrive_file_bytes(file_id):
    file_stream = download_gdrive_file_to_bytes(file_id)
    return file_stream.getvalue()

def upload_or_update_gdrive_file(folder_id, file_name, file_bytes, existing_file_id=None):
    file_name_str = str(file_name).lower()
    if file_name_str.endswith('.xlsm'):
        mime_type = 'application/vnd.ms-excel.sheet.macroEnabled.12'
    elif file_name_str.endswith('.xls'):
        mime_type = 'application/vnd.ms-excel'
    else:
        mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        
    media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type, resumable=True)
    
    if existing_file_id:
        service.files().update(fileId=existing_file_id, media_body=media, supportsAllDrives=True).execute()
        return existing_file_id
    else:
        st.error(f"❌ 拒絕建立新檔案【{file_name}】！為避免 Google 空間配額與權限錯誤，請先手動於雲端建立該檔案。")
        st.stop()

def format_gdrive_time(time_str):
    if not time_str: return "❌ 雲端檔案尚未建立/不存在"
    try:
        dt = datetime.datetime.strptime(time_str.split('.')[0], "%Y-%m-%dT%H:%M:%S")
        dt = dt + datetime.timedelta(hours=8)
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    except:
        return time_str

# =========================================================================
# 🧹 3. 資料清洗與輔助工具
# =========================================================================
def calculate_md5(file_bytes):
    return hashlib.md5(file_bytes).hexdigest()

def clean_barcode(val):
    if pd.isna(val): return ""
    s = str(val).strip().replace(" ", "").replace("\xa0", "")
    if s.lower() == "nan" or s == "": return ""
    if s.endswith(".0"): s = s[:-2]
    if "e+" in s.lower() or "e" in s.lower():
        try: s = f"{float(s):.0f}"
        except: pass
    return s

def process_smart_headers(df_raw, header_row_idx):
    if header_row_idx >= 5 and len(df_raw) > 5:
        row_5 = df_raw.iloc[4].fillna("").astype(str).str.strip()
        row_6 = df_raw.iloc[5].fillna("").astype(str).str.strip()
        new_cols = []
        for idx in range(len(df_raw.columns)):
            c5 = row_5.iloc[idx] if idx < len(row_5) else ""
            c6 = row_6.iloc[idx] if idx < len(row_6) else ""
            if any(k in c5 for k in ["訂購", "CTN", "內盒", "數量", "單價"]) and c5 != c6 and c6 != "":
                new_cols.append(f"{c5}{c6}")
            elif c6 != "": new_cols.append(c6)
            else: new_cols.append(c5)
        return new_cols
    else:
        return [str(col).strip() for col in df_raw.iloc[header_row_idx].fillna("")]