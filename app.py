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
# 🛠️ check python-calamine
# =========================================================================
try:
    import calamine
    HAS_CALAMINE = True
except ImportError:
    HAS_CALAMINE = False

# =========================================================================
# 🌐 1. Google Drive 雲端資料夾 ID 定義與權限初始化
# =========================================================================
ID_PROD_FOLDER = "1NtMAYb-SvdH6XMmqB5G07ttB-NuWCqDV"          # 商品列表 資料夾 ID
ID_PRICE_SUMMARY_FOLDER = "1ZM4MscX0UO6rUHjKv-mN5fKDwxg53maZ" # 價格統整表 資料夾 ID
ID_SHOPEE_FOLDER = "17eiGnXyU4KwNS6IR5bubBPti46SKXMH0"        # 蝦皮商品清單 資料夾 ID
ID_HISTORY_INWARD_FOLDER = "1ZQ7x4BdRc6BJlURxQ61JqDKrKF7h_vSH"# 歷史入庫單 資料夾 ID
ID_BASE_FOLDER = "1HjMt8z8DXlqGhSqe50_hDR3f4LpVLK_w"          # 麗嬰採購統整 資料夾 ID

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

service = init_drive_service()

# =========================================================================
# 🔍 2. 雲端核心實戰工具與搜尋常式
# =========================================================================
@st.cache_data(ttl=3600)  # 優化：快取雲端檔案搜尋結果 1 小時
def get_cached_gdrive_id(folder_id, file_name_keyword):
    """在指定的 Google Drive 資料夾中，根據名稱搜尋檔案並返回其 ID、修改時間、完整檔名"""
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
    """列出雲端資料夾內的所有 Excel 檔案"""
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
    """將雲端檔案下載至記憶體中"""
    request = service.files().get_media(fileId=file_id)
    file_stream = io.BytesIO()
    downloader = MediaIoBaseDownload(file_stream, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
    file_stream.seek(0)
    return file_stream

def upload_or_update_gdrive_file(folder_id, file_name, file_bytes, existing_file_id=None):
    """【強制覆寫優化版】一律不允許機器人 Create 新檔案，強制執行 Update 覆寫"""
    media = MediaIoBaseUpload(
        io.BytesIO(file_bytes), 
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
        resumable=True
    )
    
    if existing_file_id:
        # 🟢 只准執行 Update，並加入雙重防呆參數
        service.files().update(
            fileId=existing_file_id, 
            media_body=media, 
            supportsAllDrives=True
        ).execute()
        return existing_file_id
    else:
        # 🔴 防呆攔截：阻斷任何可能引發 0GB 空間配額爆炸的 create 行為
        st.error(f"❌ 拒絕建立新檔案【{file_name}】！為避免 Google 空間配額錯誤，請先手動於雲端建立該空白檔案，並將 ID 配置於系統中。")
        st.stop()

def format_gdrive_time(time_str):
    if not time_str:
        return "❌ 雲端檔案尚未建立/不存在"
    try:
        dt = datetime.datetime.strptime(time_str.split('.')[0], "%Y-%m-%dT%H:%M:%S")
        dt = dt + datetime.timedelta(hours=8)
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    except:
        return time_str

def calculate_md5(file_bytes):
    return hashlib.md5(file_bytes).hexdigest()

# -------------------------------------------------------------------------
# 動態偵測各核心主表的雲端 ID ＆ 最後修改時間 (使用 Cache 加速)
# -------------------------------------------------------------------------
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

# ==========================================
# 🧭 5. 側邊欄：導覽控制台
# ==========================================
st.sidebar.markdown("## 🏢 進銷存中央管理系統")
st.sidebar.write("---")

main_module = st.sidebar.selectbox(
    "🎯 請選擇核心管理模組：",
    ["🏪 sitegiant 電商整合管理", "📦 商品蝦皮麗嬰統整管理"]
)
st.sidebar.write("") 

if "商品蝦皮麗嬰統整管理" in main_module:
    st.sidebar.markdown("### 🛠️ 整合合併轉換功能")
    sub_page = st.sidebar.radio(
        "請選擇執行項目：",
        ["🧠 PowerQuery 執行三表整合", "⚖️ 麗嬰商品表合併和與審核", "📈 蝦皮商品清單轉換", "📊 PowerQuery 三表整合歷史紀錄", "🔍 商品清單紀錄查詢", "🔍 麗嬰商品總表數據查詢"],
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
# 🖥️ 6. 各子分頁功能邏輯處理
# ==========================================
st.title(f"{sub_page}")
st.info(f"目前導覽路徑： {main_module} ➔ {sub_page}")
st.write("---")

# ==========================================
# 🖥️ 7. 全域獨立功能: 三表 PowerQuery 整合、計算財務指標
# ==========================================
def run_powerquery_and_update_gdrive():
    """
    全域獨立功能：執行三表 PowerQuery 整合、計算財務指標，並強制 Update 覆寫至雲端統整表。
    可在任何分頁直接調用，成功回傳 True，失敗回傳 False。
    """
    if not (ID_LOCAL_PROD and ID_SHOPEE_MASTER and ID_MASTER_FILE):
        st.error("❌ 缺少核心資料主檔案 ID，無法啟動三表整合！")
        return False
        
    try:
        # 1. 跨表讀取數據
        engine_kw = {"engine": "calamine"} if HAS_CALAMINE else {}
        df_liying = pd.read_excel(download_gdrive_file_to_bytes(ID_MASTER_FILE), sheet_name="麗嬰國際產品總表", **engine_kw)
        df_p = pd.read_excel(download_gdrive_file_to_bytes(ID_LOCAL_PROD), sheet_name="商品iSKU清單", **engine_kw)
        df_s = pd.read_excel(download_gdrive_file_to_bytes(ID_SHOPEE_MASTER), sheet_name="蝦皮商品列表", **engine_kw)
        
        # 2. 資料清洗與標準化
        df_liying['條碼'] = df_liying['條碼'].astype(str).str.strip().str.split('.').str[0]
        df_p["自定義編碼"] = df_p["自定義編碼"].astype(str).str.strip().str.split('.').str[0]
        df_s["iSKU"] = df_s["iSKU"].astype(str).str.strip().str.split('.').str[0]
        
        # 3. 模擬 PowerQuery 進行多表 Merge 關聯
        df_merge1 = pd.merge(df_p, df_s[["iSKU", "GTIN", "價格"]], left_on="自定義編碼", right_on="iSKU", how="left").rename(columns={"GTIN": "蝦皮GTIN", "價格": "蝦皮售價"})
        df_merge1["c"] = df_merge1["c"].astype(str).str.strip().str.split('.').str[0]
        
        df_final = pd.merge(df_merge1, df_liying[["條碼", "零售價", "含稅"]], left_on="c", right_on="條碼", how="left").rename(columns={"零售價": "麗嬰零售價", "含稅": "麗嬰批發含稅價", "條碼": "麗嬰條碼"})
        df_final["麗嬰商品"] = df_final["麗嬰條碼"].apply(lambda x: None if pd.isna(x) else "v")
        
        # 4. 財務與稅款指標動態計算
        for c in ["蝦皮售價", "麗嬰零售價", "麗嬰批發含稅價"]:
            df_final[c] = pd.to_numeric(df_final[c], errors='coerce')
            
        df_final["麗嬰零售八折"] = df_final["麗嬰零售價"] * 0.8
        df_final["麗嬰八折比蝦皮貴"] = df_final.apply(lambda r: "v" if (pd.notna(r["麗嬰零售八折"]) and pd.notna(r["蝦皮售價"]) and r["麗嬰零售八折"] > r["蝦皮售價"]) else None, axis=1)
        df_final["麗嬰未稅價"] = df_final["麗嬰批發含稅價"].apply(lambda x: round(x / 1.05) if pd.notna(x) else None)
        df_final["麗嬰稅款"] = df_final.apply(lambda r: round(r["麗嬰批發含稅價"] - r["麗嬰未稅價"]) if (pd.notna(r["麗嬰批發含稅價"]) and pd.notna(r["麗嬰未稅價"])) else None, axis=1)
        
        # 清除不必要的欄位並寫入全域狀態（供預覽與下載使用）
        df_pq_final = df_final.drop(columns=["iSKU"], errors="ignore")
        st.session_state['pq_result'] = df_pq_final
        
        # 5. 將結果轉為記憶體二進位流並準備覆寫雲端
        output_stream = io.BytesIO()
        with pd.ExcelWriter(output_stream, engine='openpyxl') as writer:
            df_pq_final.to_excel(writer, index=False, sheet_name="商品蝦皮麗嬰價格統整表")
        output_stream.seek(0)
        
        # 尋找雲端現有的「商品蝦皮麗嬰價格統整表」檔案 ID
        existing_summary_id, _, _ = get_cached_gdrive_id(ID_PRICE_SUMMARY_FOLDER, "商品蝦皮麗嬰價格統整表")
        
        if existing_summary_id:
            # 💡 呼叫你修改過、具備安全 Update 的全域函式進行覆寫
            upload_or_update_gdrive_file(
                folder_id=ID_PRICE_SUMMARY_FOLDER, 
                file_name="商品蝦皮麗嬰價格統整表.xlsx", 
                file_bytes=output_stream.getvalue(), 
                existing_file_id=existing_summary_id
            )
            
            # 6. 強制剔除本地快取，確保其他分頁刷新時能立即抓到最新數據
            if "gdrive_id_cache" in st.session_state:
                cache_key = f"{ID_PRICE_SUMMARY_FOLDER}_商品蝦皮麗嬰價格統整表"
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
# 子功能 1：📊 三表整合歷史
# -------------------------------------------------------------------------
if sub_page == "📊 PowerQuery 三表整合歷史紀錄":
    st.subheader("🔄 三表整合歷史紀錄追蹤")
    hist_pq_files = list_gdrive_files(ID_PRICE_SUMMARY_FOLDER)
    if not hist_pq_files:
        st.warning(f"💡 提示：目前雲端資料夾內尚無任何歷史檔案，請至『🧠 PowerQuery 三表整合』執行新建轉換。")
    else:
        file_options = {f['name']: f['id'] for f in hist_pq_files}
        selected_pq_file = st.selectbox("🎯 請選擇欲調閱的歷史整合報告：", list(file_options.keys()))
        
        if selected_pq_file:
            try:
                target_id = file_options[selected_pq_file]
                file_bytes = download_gdrive_file_to_bytes(target_id)
                # 優化：純讀取歷史紀錄，嘗試使用 calamine 加速
                df_pq_view = pd.read_excel(file_bytes, engine="calamine" if HAS_CALAMINE else None)
                st.markdown(f"📄 **目前調閱雲端檔案**：`{selected_pq_file}` ｜ 📊 **資料總項數**：`{len(df_pq_view)} 筆`")
                st.dataframe(df_pq_view, use_container_width=True)
                st.download_button(label="🔄 重新下載此歷史整合表 (.xlsx)", data=file_bytes.getvalue(), file_name=selected_pq_file, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"❌ 讀取雲端備份檔案失敗: {str(e)}")

# -------------------------------------------------------------------------
# 子功能 2：🧠 PowerQuery 三表整合 (加強下載報表與顯示雲端統整表最後修改時間)
# -------------------------------------------------------------------------
elif sub_page == "🧠 PowerQuery 執行三表整合":
    st.subheader("🔍 三表數據追蹤")
    
    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("📦 商品列表 (商品iSKU清單)", "已對接" if ID_LOCAL_PROD else "❌ 未偵測到")
        st.caption(f"📅 最後修改時間: \n`{format_gdrive_time(TIME_PROD)}`")
    with c2:
        st.metric("🧡 蝦皮資料庫主表", "已對接" if ID_SHOPEE_MASTER else "❌ 未偵測到")
        st.caption(f"📅 最後修改時間: \n`{format_gdrive_time(TIME_SHOPEE)}`")
    with c3:
        st.metric("🧸 麗嬰產品總表", "已對接" if ID_MASTER_FILE else "❌ 未偵測到")
        st.caption(f"📅 最後修改時間: \n`{format_gdrive_time(TIME_MASTER)}`")

    st.write("---")

    # 🌟 全新加入：在執行前先到雲端抓取「商品蝦皮麗嬰價格統整表」的目前狀態與最後修改時間
    st.subheader("📊 雲端『商品蝦皮麗嬰價格統整表』當前狀態")
    existing_summary_id, existing_summary_time, _ = get_cached_gdrive_id(ID_PRICE_SUMMARY_FOLDER, "商品蝦皮麗嬰價格統整表")
    
    if existing_summary_id:
        st.info(f"🟢 雲端已存在統整表檔案 ｜ 📅 最後修改時間：`{format_gdrive_time(existing_summary_time)}`")
    else:
        st.warning("⚠️ 雲端目前尚未建立『商品蝦皮麗嬰價格統整表』，回寫時系統將會自動全新建立。")

    st.write("---")

    if st.button("🛠️ 啟動三表整合與財務指標計算", type="primary", use_container_width=True):
        if not (ID_LOCAL_PROD and ID_SHOPEE_MASTER and ID_MASTER_FILE):
            st.error("❌ 無法啟動三表整合！請確認雲端對應資料夾內是否缺少必要的核心資料主檔案。")
        else:
            with st.spinner("正在由雲端載入數據流並進行大數據跨表計算..."):
                try:
                    # 優化：跨表巨量讀取，指定使用 calamine 引擎大幅加速
                    engine_kw = {"engine": "calamine"} if HAS_CALAMINE else {}
                    df_liying = pd.read_excel(download_gdrive_file_to_bytes(ID_MASTER_FILE), sheet_name="麗嬰國際產品總表", **engine_kw)
                    df_p = pd.read_excel(download_gdrive_file_to_bytes(ID_LOCAL_PROD), sheet_name="商品iSKU清單", **engine_kw)
                    df_s = pd.read_excel(download_gdrive_file_to_bytes(ID_SHOPEE_MASTER), sheet_name="蝦皮商品列表", **engine_kw)
                    
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
                    
                    # 將結果存入 pq_result 機制
                    st.session_state['pq_result'] = df_final.drop(columns=["iSKU"], errors="ignore")
                    st.success("🎉 三表 PowerQuery 交叉聯結與財務指標計算整合完成！")
                except Exception as e:
                    st.error(f"❌ 錯誤: {str(e)}")

    # ── 當有整合結果存在時，顯示報表預覽、本地下載功能、以及雲端回寫機制 ──
    if 'pq_result' in st.session_state and st.session_state['pq_result'] is not None:
        df_result = st.session_state['pq_result']
        
        st.subheader("📋 整合聯結情報報表輸出預覽")
        st.markdown(f"📊 **目前整合結果資料總項數**：`{len(df_result)} 筆`")
        st.dataframe(df_result, use_container_width=True)
        
        # 建立下載與回寫的功能按鈕排版 (左右雙欄對齊)
        col_btn1, col_btn2 = st.columns(2)
        
        with col_btn1:
            # ✨ 功能 1：下載整合報表到本機電腦
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
            # ✨ 功能 2：強制覆寫雲端統整表
            with col_btn2:
                if st.button("🔄 執行：將整合結果回寫並更新至雲端『商品蝦皮麗嬰價格統整表』", type="secondary", use_container_width=True):
                    with st.spinner("💾 正在覆寫更新雲端現有統整表檔案..."):
                        # ── 🌟 直接一鍵調用全域 Function ──
                        if run_powerquery_and_update_gdrive():
                            st.success("✅ 雲端統整表已成功同步覆寫更新！")
                            st.info("💡 重新整理頁面後，上方將會顯示最新的修改時間。")
                        
# -------------------------------------------------------------------------
# 子功能 3：🔍 麗嬰商品總表數據查詢 (新增多筆條碼批次查詢功能)
# -------------------------------------------------------------------------
elif sub_page == "🔍 麗嬰商品總表數據查詢":
    st.subheader("📋 麗嬰採購產品總表資料庫分頁動態檢視")
    
    # 延遲載入 metadata，避免全域耗能
    _, _, _, _, all_sheets, _ = load_master_data(ID_MASTER_FILE)
    
    if all_sheets:
        view_sheets = [s for s in all_sheets if s != "麗嬰產品新採購單"]
        selected_sheet = st.selectbox("請選擇數據分頁：", view_sheets)
        
        # ── 🌟 新增功能：切換查詢模式 ──
        search_mode = st.radio("🎯 請選擇查詢模式：", ["多筆條碼價格查詢", "模糊關鍵字搜尋"], horizontal=True)
        
        try:
            # 使用 calamine 快速讀取預覽
            engine_kw = {"engine": "calamine"} if HAS_CALAMINE else {}
            df_view = pd.read_excel(download_gdrive_file_to_bytes(ID_MASTER_FILE), selected_sheet, **engine_kw)
            
            # 確保條碼欄位格式乾淨（去除小數點與空白）
            if "條碼" in df_view.columns:
                df_view['條碼'] = df_view['條碼'].astype(str).str.strip().str.split('.').str[0]
            
            # ── 模式 1：原本的模糊關鍵字搜尋 ──
            if search_mode == "模糊關鍵字搜尋":
                search_term = st.text_input("🔍 快速搜尋關鍵字 (支援條碼、品名、貨號模糊比對)：", placeholder="輸入搜尋內容...")
                st.metric(label=f"📊 【{selected_sheet}】當前總資料筆數", value=f"{len(df_view)} 筆")
                
                if search_term:
                    search_mask = df_view.astype(str).apply(lambda x: x.str.contains(search_term, case=False, na=False)).any(axis=1)
                    st.dataframe(df_view[search_mask], use_container_width=True)
                else:
                    st.dataframe(df_view, use_container_width=True)
            
            # ── 模式 2：🚀 全新功能 - 多筆條碼價格查詢 ──
            elif search_mode == "多筆條碼價格查詢":
                if "條碼" not in df_view.columns:
                    st.warning(f"⚠️ 當前選擇的分頁 【{selected_sheet}】 內部不含「條碼」欄位，無法使用此查詢模式。")
                else:
                    barcode_paste = st.text_area(
                        "📋 請貼上多筆國際條碼 (支援從 Excel 複製直接貼上，或以換行、空格、逗號隔開)：",
                        height=120,
                        placeholder="範例：\n4711234567890\n4711234567891"
                    )
                    
                    if barcode_paste.strip():
                        # 解析使用者輸入的條碼列表（清洗空白、換行、逗號）
                        raw_barcodes = barcode_paste.replace(',', ' ').replace('\t', ' ').split()
                        cleaned_barcodes = [str(b).strip().split('.')[0] for b in raw_barcodes if b.strip()]
                        
                        if cleaned_barcodes:
                            # 使用 Pandas isin 進行高效比對
                            df_result = df_view[df_view['條碼'].isin(cleaned_barcodes)].copy()
                            
                            # 動態回報搜尋統計
                            found_count = len(df_result)
                            input_count = len(set(cleaned_barcodes)) # 去重後的輸入條碼數
                            
                            st.success(f"🔍 查詢完畢！您輸入了 {input_count} 筆不重複條碼，成功比對出 {found_count} 筆商品資料。")
                            
                            # 挑選核心售價相關欄位置前顯示（如果欄位存在的話）
                            important_cols = ["條碼", "名稱", "零售價", "含稅"]
                            display_cols = [c for c in important_cols if c in df_result.columns] + \
                                           [c for c in df_result.columns if c not in important_cols]
                            
                            st.dataframe(df_result[display_cols], use_container_width=True)
                            
                            # 提供獨立的查詢結果下載按鈕
                            towrite_query = io.BytesIO()
                            with pd.ExcelWriter(towrite_query, engine='openpyxl') as writer:
                                df_result[display_cols].to_excel(writer, index=False, sheet_name="條碼查詢結果")
                            
                            st.download_button(
                                label="📥 下載本次條碼查詢結果報表 (.xlsx)",
                                data=towrite_query.getvalue(),
                                file_name=f"條碼批次查詢結果_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        else:
                            st.info("💡 請輸入有效的條碼。")
                    else:
                        st.info("💡 請在上方文字方塊中貼上欲查詢的條碼。")
                        
        except Exception as e:
            st.error(f"❌ 讀取分頁數據失敗: {str(e)}")

# -------------------------------------------------------------------------
# 子功能 4：⚖️ 麗嬰商品表合併和與審核 (新增歸檔後一鍵整合回寫快捷鍵)
# -------------------------------------------------------------------------
elif sub_page == "⚖️ 麗嬰商品表合併和與審核":
    st.subheader("🧸 麗嬰採購單一鍵導入與審核系統")
    
    # 延遲載入並抓取已快取的主表狀態
    df_total, df_history, df_delete_log, df_meta, _, current_max_uid = load_master_data(ID_MASTER_FILE)
    if df_total is None: st.stop()

    # 顯示歸檔成功的重大提示與快捷功能
    if 'merge_success_msg' in st.session_state:
        st.success(st.session_state['merge_success_msg'])
        
        # ── 🌟 全新亮點：歸檔成功後原地跳出快捷控制面板 ──
        st.markdown("### ⚡ 歸檔後後續自動化推薦操作")
        st.info("💡 採購單已成功存入麗嬰總表！現在您可以直接點擊下方按鈕，原地啟動跨表 PowerQuery 整合並自動更新雲端統整表。")
        
        if st.button("🚀 馬上執行：三表資料整合並自動回寫更新至雲端『商品蝦皮麗嬰價格統整表』", type="primary", use_container_width=True):
            with st.spinner("⏳ 正在跨資料庫調閱核心數據、執行大數據 VLOOKUP 計算並回寫雲端..."):
                # ── 🌟 直接一鍵調用全域 Function ──
                if run_powerquery_and_update_gdrive():
                    st.success("🎯 狂賀！三表整合計算完成，且『商品蝦皮麗嬰價格統整表』已在雲端同步覆寫更新完畢！")
                    del st.session_state['merge_success_msg']
                    
        st.write("---")

    uploaded_files = st.file_uploader("📥 選擇採購單 Excel (可多選批次上傳)", type=["xlsx", "xls", "xlsm"], accept_multiple_files=True, key="main_merge_files")
    
    if uploaded_files:
        if st.button("🚀 開始一鍵合併歸檔", type="primary"):
            success_count = dup_count = no_barcode_count = 0
            new_rows, history_records = [], []
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
                        match_found = False
                        need_insert = False
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
                    history_records.append({"檔案名稱": file.name, "md5": file_md5, "匯入時間": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
                    success_count += 1
                except Exception as e:
                    st.error(f"❌ 解析失敗: {str(e)}")
                    
            if new_rows: df_total = pd.concat([df_total, pd.DataFrame(new_rows)], ignore_index=True)
            if history_records: df_history = pd.concat([df_history, pd.DataFrame(history_records)], ignore_index=True)
            
            df_meta.iloc[0, 0] = current_max_uid
            df_total = run_cross_matching(df_total)
            
            if save_to_master_xlsm({"麗嬰國際產品總表": df_total, "已匯入採購單": df_history, "metadata": df_meta}):
                load_master_data.clear()
                get_cached_gdrive_id.clear()
                # 💡 關鍵：設定 Session 訊息後刷新頁面，下次載入時就會立刻在最上方呈現快捷按鈕！
                st.session_state['merge_success_msg'] = f"🎉 成功完成歸檔與雲端同步寫入！順利處理 {success_count} 份檔案。 (跳過重複檔: {dup_count})"
                st.rerun()

    st.write("---")
    st.subheader("⚠️ 條碼重複與衝突即時審核控制台")
    # ... 審核控制台後續邏輯完全不變 ...
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
                elif "零售價different" in remark or ("零售價不同" in remark and not barcode.startswith("🟢")): row['條碼'] = f"🟢 {barcode}"
                return row
            df_anomaly = df_anomaly.apply(inject_emoji_alerts, axis=1)
            
            st.warning("下方商品為系統抓出之條碼重複資料：🔴 代表名稱不一致，🟢 代表售價不一致。您可以直接在下方『備註』欄雙擊文字編寫 Note 紀錄！")
            
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
                    
                    if save_to_master_xlsm({"麗嬰國際產品總表": df_remaining, "刪除紀錄": df_delete_log}):
                        load_master_data.clear() # 優化：更新後清空快取
                        get_cached_gdrive_id.clear()
                        st.session_state['merge_success_msg'] = "🧹 移轉封存與自訂 Note 備註已完整同步寫入雲端主資料庫！"
                        st.rerun()
                else:
                    df_remaining_only = run_cross_matching(df_total.drop(columns=['move'], errors='ignore'))
                    if save_to_master_xlsm({"麗嬰國際產品總表": df_remaining_only, "刪除紀錄": df_delete_log}):
                        load_master_data.clear() # 優化：更新後清空快取
                        get_cached_gdrive_id.clear()
                        st.session_state['merge_success_msg'] = "📝 自訂 Note 備註內容已順利同步更新至雲端主資料庫！"
                        st.rerun()
        else:
            st.success("🟢 當前總表中沒有任何重複商品的衝突。")

# -------------------------------------------------------------------------
# 子功能 5：📈 蝦皮商品清單轉換
# -------------------------------------------------------------------------
elif sub_page == "📈 蝦皮商品清單轉換":
    st.subheader("🛍️ 蝦皮賣場商品列表iSKU結構校正")
    
    # 延遲載入蝦皮快取
    df_shopee_history, df_shopee_current_list = load_shopee_data(ID_SHOPEE_MASTER)

    uploaded_shopee = st.file_uploader("📥 上傳新的蝦皮商品清單原始報表 (.xlsx) 進行格式校正：", type=["xlsx"], key="main_shopee_upload")
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
                                     
                    out_buf_sp = io.BytesIO()
                    df_final_clean.to_excel(out_buf_sp, index=False)
                    
                    # ── 🟢 修正：強制只更新現有的「蝦皮賣場商品列表.xlsm」主表 ──
                    if ID_SHOPEE_MASTER:
                        upload_or_update_gdrive_file(ID_SHOPEE_FOLDER, "蝦皮賣場商品列表.xlsm", out_buf_sp.getvalue(), existing_file_id=ID_SHOPEE_MASTER)
                    else:
                        st.error("❌ 雲端找不到變數 `ID_SHOPEE_MASTER` 對應的現有主檔案 ID，無法執行覆寫更新。")
                        st.stop()
                    
                    new_hist_log = pd.DataFrame([{"檔案名稱": uploaded_shopee.name, "md5": shopee_md5, "匯入時間": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}])
                    df_shopee_history = pd.concat([df_shopee_history, new_hist_log], ignore_index=True)
                    
                    if save_to_shopee_master_xlsm({"蝦皮商品列表": df_final_clean, "匯入檔案": df_shopee_history}):
                        load_shopee_data.clear() 
                        get_cached_gdrive_id.clear()
                        st.session_state['shopee_clean'] = df_final_clean
                        st.success(f"🎉 蝦皮賣場商品列表iSKU結構校正完成！\n🟢 雲端現有主表 `蝦皮賣場商品列表.xlsm` 已成功同步覆寫更新！")
                        
                        # ── 🌟 新增：讓蝦皮校正完也能原地一鍵觸發更新統整表 ──
                        st.markdown("---")
                        st.markdown("### ⚡ 後續自動化推薦操作")
                        if st.button("🚀 馬上更新：以最新校正的蝦皮資料重新執行三表整合並回寫雲端", type="primary", use_container_width=True):
                            with st.spinner("⏳ 正在重新整理跨表聯結數據並回寫..."):
                                if run_powerquery_and_update_gdrive():
                                    st.success("✅ 成功！雲端『商品蝦皮麗嬰價格統整表』已同步使用最新校正後的蝦皮資料覆寫更新！")
                except Exception as e:
                    st.error(f"讀取或清洗蝦皮檔案失敗: {str(e)}")

    if 'shopee_clean' in st.session_state:
        st.dataframe(st.session_state['shopee_clean'], use_container_width=True)
        towrite_shopee = io.BytesIO()
        st.session_state['shopee_clean'].to_excel(towrite_shopee, index=False)
        st.download_button(label="📥 下載此次iSKU校正蝦皮報表 (.xlsx)", data=towrite_shopee.getvalue(), file_name=f"蝦皮清洗完成對齊表_{datetime.date.today().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# -------------------------------------------------------------------------
# 子功能 6：🔀 sitegiant 採購入庫單轉換
# -------------------------------------------------------------------------
elif sub_page == "🔀 sitegiant 採購入庫單格式轉換":
    st.subheader("🛍️ SiteGiant 採購入庫單內容填寫")
    
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
        recv_date = st.date_input("📅 選擇收貨日：", value=datetime.date.today())
        
    st.write("---")
    st.subheader("🚀 條碼與數量批次快速貼上")
    bulk_paste_area = st.text_area("請從excel複製條碼、數量 (Tab鍵或空白鍵隔開)：", height=120, placeholder="條碼 數量")
    
    if bulk_paste_area.strip():
        lines = [line.strip() for line in bulk_paste_area.strip().split('\n') if line.strip()]
        parsed_rows = []
        for line in lines:
            tokens = line.split(',') if "," in line else (line.split('\t') if "\t" in line else line.split())
            if tokens:
                b_code = str(tokens[0]).strip().split('.')[0]
                q_val = 1 
                if len(tokens) > 1:
                    try: q_val = int(float(str(tokens[1]).strip()))
                    except: q_val = 1
                parsed_rows.append({"國際條碼": b_code, "數量": q_val})
        if parsed_rows:
            st.session_state['inward_input_df'] = pd.DataFrame(parsed_rows)
            st.toast("重新載入成功！已同步填入下方明細表格。", icon="🚀")

    st.write("---")
    input_df = st.data_editor(st.session_state['inward_input_df'], num_rows="dynamic", use_container_width=True, key="inward_grid")
    st.session_state['inward_input_df'] = input_df 
    
    if st.button("✨ 執行貨品名稱和成本稅款導入", type="primary", use_container_width=True):
        if not order_no.strip() or input_df.empty: 
            st.error("❌ 轉換失敗！請填入銷貨單號與有效明細。")
        else:
            try:
                if ID_PRICE_SUMMARY:
                    engine_kw = {"engine": "calamine"} if HAS_CALAMINE else {}
                    df_ref = pd.read_excel(download_gdrive_file_to_bytes(ID_PRICE_SUMMARY), **engine_kw)
                    df_ref['c_clean'] = df_ref['c'].astype(str).str.strip().str.split('.').str[0]
                        
                    result_rows = []
                    for row in input_df.itertuples(index=False):
                        barcode_input = str(row.國際條碼).strip().split('.')[0] if pd.notna(row.國際條碼) else ""
                        if barcode_input in ["", "0", "nan", "None"]: continue
                        qty = int(row.數量) if pd.notna(row.數量) else 0
                    
                        sku_final = "⚠️ 提示：須新增iSKU"
                        prod_name = "請確認商品列表和統整表是否已經更新"
                        category = ""
                        keywords = ""
                        # ── ⚙️ 核心邏輯：預設為 None，讓畫面呈現乾淨空白 ──
                        cost_val = None 
                        tax_val = None
                    
                        if not df_ref.empty and 'c_clean' in df_ref.columns:
                            match = df_ref[df_ref['c_clean'] == barcode_input]
                            if not match.empty:
                                match_row = match.iloc[0]
                                prod_name = match_row.get('品名', match_row.get('名稱', ''))
                                sku = match_row.get('自定義編碼', '')
                                sku_final = sku if pd.notna(sku) and str(sku).strip() != "" else "⚠️ 提示：須新增iSKU"
                                category = match_row.get('分類定義', '')
                                keywords = match_row.get('產品關鍵字', '')
                            
                                # ── ⚙️ 廠商判定邏輯：只有麗嬰才從總表抓成本與稅款，其餘維持 None ──
                                if vendor_name == "麗嬰":
                                    or_raw = match_row.get('麗嬰零售價', None)
                                    sal_raw = match_row.get('麗嬰批發含稅價',None)
                                    c_raw = match_row.get('麗嬰未稅價', None)
                                    t_raw = match_row.get('麗嬰稅款', None)
                                    cost_val = float(c_raw) if pd.notna(c_raw) else None
                                    tax_val = float(t_raw) if pd.notna(t_raw) else None
                                    or_val = float(or_raw) if pd.notna(or_raw) else None
                                    sal_val = float(sal_raw) if pd.notna(sal_raw) else None

                        result_rows.append({
                            "收貨日": str(recv_date), "國際條碼": barcode_input,
                            "庫存SKU": sku_final, "庫存貨品名稱": prod_name, 
                            "麗嬰零售價": or_val, "麗嬰批發含稅價": sal_val,
                            "成本": cost_val, "稅款": tax_val, "數量": qty,
                            "分類定義": category, "產品關鍵字": keywords
                        })                    
                else:
                    st.error("❌ 雲端找不到『商品蝦皮麗嬰價格統整表』。")
                    
                if result_rows:
                    st.session_state['inward_result_df'] = pd.DataFrame(result_rows)
                    st.session_state['current_vendor_name'] = vendor_name
                    st.session_state['current_order_no'] = order_no
                    st.success(f"🚀 格式勾稽完成！廠商已設定為：【{vendor_name}】")
                
            except Exception as e: 
                st.error(f"❌ 錯誤: {str(e)}")

    # ── 📝 步驟 2：表格即時可編輯渲染與動態加總區塊 ──
    if 'inward_result_df' in st.session_state:
        res_df = st.session_state['inward_result_df']
        current_vendor = st.session_state.get('current_vendor_name', '未命名廠商')
        current_order = st.session_state.get('current_order_no', '0000')
        
        target_columns = ["國際條碼","庫存SKU", "庫存貨品名稱", "麗嬰零售價", "麗嬰批發含稅價", "成本", "稅款", "數量"]
        available_cols = [col for col in target_columns if col in res_df.columns]
        df_download = res_df[available_cols].copy()
        
        st.markdown(f"📋 **【{current_vendor}】入庫明細編輯與預覽（請雙擊「成本」或「稅款」直接修改數字）：**")
        
        # 使用 data_editor 渲染，強制定義為純數值欄位（允許 None 留白編輯）
        edited_inward_df = st.data_editor(
            df_download,
            use_container_width=True,
            disabled=["國際條碼","庫存SKU", "庫存貨品名稱", "數量"],
            column_config={
                "成本": st.column_config.NumberColumn(
                    "成本", help="請手動輸入未稅成本", min_value=0.0, format="%.2f"
                ),
                "稅款": st.column_config.NumberColumn(
                    "稅款", help="請手動輸入營業稅款", min_value=0.0, format="%.2f"
                ),
            },
            key="inward_items_editor_final"
        )
        
        # ── 📊 步驟 3：【核心動態即時加總】防呆過濾 None 值 ──
        h_cost = 0.0
        h_tax = 0.0
        for h_row in edited_inward_df.itertuples(index=False):
            h_qty = int(h_row.數量) if (hasattr(h_row, '數量') and pd.notna(h_row.數量)) else 0
            
            # 動態累加成本 (成本 * 數量)
            if hasattr(h_row, '成本') and pd.notna(h_row.成本):
                try: h_cost += float(h_row.成本) * h_qty
                except: pass
                
            # 動態累加稅款 (稅款 * 數量)
            if hasattr(h_row, '稅款') and pd.notna(h_row.稅款):
                try: h_tax += float(h_row.稅款) * h_qty
                except: pass
        
        # ── 🖥️ 步驟 4：動態渲染金額看板 ──
        st.write("---")
        st.markdown("#### 📊 本張單據入庫成本稅款即時統計看板 (隨手動輸入動態更新)")
        c_tot1, c_tot2 = st.columns(2)
        with c_tot1: 
            st.metric(label="💰 當前單據成本未稅總金額 (成本 * 數量)", value=f"$ {h_cost:,.2f} 元")
        with c_tot2: 
            st.metric(label="🧾 當前單據營業稅總金額 (稅款 * 數量)", value=f"$ {h_tax:,.2f} 元")
        st.write("---")
        
        # ── 💾 步驟 5：儲存與下載 ──
        towrite_inward = io.BytesIO()
        with pd.ExcelWriter(towrite_inward, engine='openpyxl') as writer:
            edited_inward_df.to_excel(writer, index=False, sheet_name="SiteGiant入庫單")
            
        final_filename = f"sitegiant採購入庫單_{recv_date}_{current_vendor}_{current_order}.xlsx"

       # upload_or_update_gdrive_file(ID_HISTORY_INWARD_FOLDER, final_filename, towrite_inward.getvalue())
        st.download_button(label=f"📥 下載sitegiant格式採購入庫單 ({final_filename})", data=towrite_inward.getvalue(), file_name=final_filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

# -------------------------------------------------------------------------
# 子功能 7：📜 sitegiant 歷史入庫單紀錄
# -------------------------------------------------------------------------
elif sub_page == "📜 sitegiant 歷史入庫單紀錄":
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
                    
                    # 動態累加成本：先確認欄位存在且不是 NaN (空白)，才進行加總
                    if hasattr(h_row, '成本') and pd.notna(h_row.成本):
                        try: 
                            h_cost += float(h_row.成本) * h_qty
                        except ValueError: 
                            pass
                            
                    # 動態累加稅款：先確認欄位存在且不是 NaN (空白)，才進行加總
                    if hasattr(h_row, '稅款') and pd.notna(h_row.稅款):
                        try: 
                            h_tax += float(h_row.稅款) * h_qty
                        except ValueError: 
                            pass
                    
                st.markdown("#### 📊 本張單據入庫成本稅款")
                c_tot1, c_tot2 = st.columns(2)
                with c_tot1: 
                    st.metric(label="💰 成本未稅總金額 (成本 * 數量)", value=f"$ {h_cost:,.2f} 元")
                with c_tot2: 
                    st.metric(label="🧾 營業稅總金額 (稅款 * 數量)", value=f"$ {h_tax:,.2f} 元")    
                
                
                st.download_button(label="🔄 下載此歷史採購入庫單", data=file_bytes.getvalue(), file_name=selected_hist_file, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e: st.error(f"❌ 讀取失敗: {str(e)}")

# -------------------------------------------------------------------------
# 子功能 8：🔍 商品清單紀錄查詢
# -------------------------------------------------------------------------
elif sub_page == "🔍 商品清單紀錄查詢":
    st.subheader("📊 歷史商品清單紀錄查詢")
    hist_files = list_gdrive_files(ID_PROD_FOLDER)
    if not hist_files: st.warning(f"💡 目前雲端無歷史單據。")
    else:
        file_options = {f['name']: f['id'] for f in hist_files}
        selected_hist_file = st.selectbox("🎯 商品清單紀錄:", list(file_options.keys()))
        if selected_hist_file:
            try:
                target_id = file_options[selected_hist_file]
                file_bytes = download_gdrive_file_to_bytes(target_id)
                df_hist_view = pd.read_excel(file_bytes, engine="calamine" if HAS_CALAMINE else None)
                st.markdown(f"📄 **當前雲端檔案**：`{selected_hist_file}` ｜ 📊 **iSKU品項數**：`{len(df_hist_view)} 筆`")
                st.dataframe(df_hist_view, use_container_width=True)
                st.download_button(label="🔄 下載此歷史商品清單", data=file_bytes.getvalue(), file_name=selected_hist_file, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e: st.error(f"❌ 讀取失敗: {str(e)}")
