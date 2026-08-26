import streamlit as st
import pandas as pd
import openpyxl
import hashlib
import datetime
import io
import os
import re
import json
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
DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive"]
XLSX_MIME_QUERY = (
    "(mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' "
    "or mimeType = 'application/vnd.ms-excel.sheet.macroEnabled.12')"
)
BATCH_EDIT_DATE_RE = re.compile(r"batch_edit_basic_info_all_(\d{2})-(\d{2})-(\d{4})", re.I)
MASS_UPDATE_DATE_RE = re.compile(r"mass_update_sales_info_3062950_(\d{8})", re.I)


def _in_streamlit():
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        return get_script_run_ctx() is not None
    except Exception:
        return False


def _notify_error(msg):
    if _in_streamlit():
        st.error(msg)
    else:
        print(msg)


def _stop_or_raise(msg):
    _notify_error(msg)
    if _in_streamlit():
        st.stop()
    raise RuntimeError(msg)


def _fix_private_key_pem(pk):
    if not isinstance(pk, str) or not pk.strip():
        return pk
    pk = pk.replace("\ufeff", "").strip()
    if (pk.startswith('"') and pk.endswith('"')) or (pk.startswith("'") and pk.endswith("'")):
        pk = pk[1:-1].strip()
    pk = pk.replace("\r\n", "\n").replace("\r", "\n")
    if "\\n" in pk:
        pk = pk.replace("\\n", "\n")
    begin_idx = pk.find("-----BEGIN ")
    if begin_idx > 0:
        pk = pk[begin_idx:]
    elif begin_idx < 0:
        end_idx = pk.find("-----END ")
        if end_idx >= 0:
            body = pk[:end_idx].strip()
            footer = pk[end_idx:].lstrip()
            pk = "-----BEGIN PRIVATE KEY-----\n" + body + "\n" + footer
        else:
            pk = "-----BEGIN PRIVATE KEY-----\n" + pk.strip() + "\n-----END PRIVATE KEY-----\n"
    if "-----BEGIN " in pk and "\n" not in pk:
        pk = re.sub(r"(-----BEGIN [A-Z0-9 ]+-----)", r"\1\n", pk)
        pk = re.sub(r"(-----END [A-Z0-9 ]+-----)", r"\n\1", pk)
    if not pk.endswith("\n"):
        pk += "\n"
    return pk


def _coerce_service_account_info(raw):
    """Accept a dict (TOML table) or a JSON string; return a dict with a framed PEM."""
    if raw is None:
        return None
    if isinstance(raw, bytes):
        raw = raw.decode("utf-8")
    if isinstance(raw, str):
        raw = raw.strip()
        if not raw:
            return None
        raw = json.loads(raw)
        if isinstance(raw, str):
            raw = json.loads(raw)
    if not isinstance(raw, dict):
        try:
            raw = dict(raw)
        except Exception:
            return None
    info = dict(raw)
    if "private_key" in info:
        info["private_key"] = _fix_private_key_pem(info.get("private_key"))
    return info


def load_google_service_account_info():
    env_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if env_json:
        return _coerce_service_account_info(env_json)
    env_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
    if env_path and os.path.isfile(env_path):
        with open(env_path, encoding="utf-8") as f:
            return _coerce_service_account_info(json.load(f))
    try:
        return _coerce_service_account_info(st.secrets["textkey"])
    except Exception:
        return None


def _build_drive_service():
    info = load_google_service_account_info()
    if not info:
        raise RuntimeError("缺少 Google 憑證：請設定 GOOGLE_SERVICE_ACCOUNT_JSON 或 Streamlit Secrets `textkey`。")
    credentials = service_account.Credentials.from_service_account_info(info, scopes=DRIVE_SCOPES)
    return build("drive", "v3", credentials=credentials)


@st.cache_resource
def init_drive_service():
    try:
        return _build_drive_service()
    except Exception as e:
        if _in_streamlit():
            st.error(f"❌ 無法讀取 Google Drive 憑證。錯誤訊息: {str(e)}")
            st.info("💡 請確認 Secrets `textkey` 或環境變數 GOOGLE_SERVICE_ACCOUNT_JSON。")
            st.stop()
        raise


def get_drive_service():
    if _in_streamlit():
        return init_drive_service()
    return _build_drive_service()

# =========================================================================
# 🔍 2. 雲端核心實戰工具與搜尋常式
# =========================================================================
def _list_gdrive_files_raw(folder_id, name_contains=None):
    query = f"'{folder_id}' in parents and {XLSX_MIME_QUERY} and trashed = false"
    if name_contains:
        query += f" and name contains '{name_contains}'"
    files = []
    page_token = None
    while True:
        results = get_drive_service().files().list(
            q=query,
            fields="nextPageToken, files(id, name, modifiedTime)",
            pageSize=100,
            pageToken=page_token,
        ).execute()
        files.extend(results.get("files", []))
        page_token = results.get("nextPageToken")
        if not page_token:
            break
    return files


@st.cache_data(ttl=3600)
def get_cached_gdrive_id(folder_id, file_name_keyword):
    try:
        files = _list_gdrive_files_raw(folder_id, name_contains=file_name_keyword)
        if files:
            files.sort(key=lambda x: x.get("modifiedTime") or "", reverse=True)
            return files[0]["id"], files[0]["modifiedTime"], files[0]["name"]
    except Exception:
        pass
    return None, None, None

def list_gdrive_files(folder_id, name_contains=None):
    try:
        files = _list_gdrive_files_raw(folder_id, name_contains=name_contains)
        files.sort(key=lambda x: x["name"], reverse=True)
        return files
    except Exception as e:
        _notify_error(f"掃描雲端資料夾失敗: {str(e)}")
        return []


def _parse_named_file_date(name, kind):
    if kind == "batch_edit":
        m = BATCH_EDIT_DATE_RE.search(name or "")
        if m:
            return datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    elif kind == "mass_update":
        m = MASS_UPDATE_DATE_RE.search(name or "")
        if m:
            return datetime.datetime.strptime(m.group(1), "%Y%m%d").date()
    return None


def pick_latest_gdrive_file(folder_id, name_contains, kind):
    """依檔名內日期取最新檔；同日再用 modifiedTime。"""
    files = list_gdrive_files(folder_id, name_contains=name_contains)
    ranked = []
    for f in files:
        file_date = _parse_named_file_date(f.get("name"), kind)
        if file_date:
            ranked.append((file_date, f.get("modifiedTime") or "", f))
    if not ranked:
        return None
    ranked.sort(key=lambda x: (x[0], x[1]), reverse=True)
    return ranked[0][2]


def download_gdrive_file_to_bytes(file_id):
    request = get_drive_service().files().get_media(fileId=file_id)
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
        get_drive_service().files().update(fileId=existing_file_id, media_body=media, supportsAllDrives=True).execute()
        return existing_file_id
    _stop_or_raise(f"❌ 拒絕建立新檔案【{file_name}】！為避免 Google 空間配額與權限錯誤，請先手動於雲端建立該檔案。")

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

# =========================================================================
# 📦 4. 資料庫共用載入常式
# =========================================================================
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


# =========================================================================
# 💾 5. 核心資料庫讀寫與勾稽常式
# =========================================================================

def save_to_master_xlsm(sheets_dict, file_id, folder_id, file_name):
    if not file_id:
        _notify_error("❌ 雲端找不到核心總表檔案")
        return False
    try:
        master_bytes = download_gdrive_file_to_bytes(file_id)
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
        upload_or_update_gdrive_file(folder_id, file_name or "麗嬰採購產品總表.xlsm", out_buf.getvalue(), existing_file_id=file_id)
        return True
    except Exception as e:
        _notify_error(f"❌ 寫入雲端資料庫發生錯誤: {str(e)}")
        return False

def save_to_shopee_master_xlsm(sheets_dict, file_id, folder_id, file_name):
    try:
        if file_id:
            shopee_bytes = download_gdrive_file_to_bytes(file_id)
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
        upload_or_update_gdrive_file(
            folder_id, 
            file_name or "蝦皮賣場商品列表.xlsm", 
            out_buf.getvalue(), 
            existing_file_id=file_id
        )
        return True
    except Exception as e:
        _notify_error(f"❌ 寫入雲端蝦皮資料庫發生錯誤: {str(e)}")
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

@st.cache_data(ttl=600)
def load_master_data(file_id):
    if not file_id: return None, None, None, None, None, 3473
    try:
        master_bytes = download_gdrive_file_to_bytes(file_id)
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
        _notify_error(f"🔴 讀取雲端主資料庫失敗。錯誤: {str(e)}")
        return None, None, None, None, None, 3473


def _file_payload_bytes(file_bytes):
    if isinstance(file_bytes, io.BytesIO):
        file_bytes.seek(0)
        data = file_bytes.getvalue()
        file_bytes.seek(0)
        return data
    return file_bytes


def load_sitegiant_batch_name_map(folder_id):
    """c/UPC -> 庫存貨品名稱，以及選到的 batch_edit 檔案 metadata。"""
    latest = pick_latest_gdrive_file(folder_id, "batch_edit_basic_info_all", "batch_edit")
    if not latest:
        return {}, None
    engine_kw = {"engine": "calamine"} if HAS_CALAMINE else {}
    df = pd.read_excel(download_gdrive_file_to_bytes(latest["id"]), dtype=str, **engine_kw)
    df.columns = df.columns.astype(str).str.strip().str.replace("\n", "")
    upc_col = next((c for c in df.columns if "國際條碼" in c or c.upper() == "UPC"), None)
    name_col = "庫存貨品名稱" if "庫存貨品名稱" in df.columns else None
    if not upc_col or not name_col:
        return {}, latest
    df["_upc"] = df[upc_col].map(clean_barcode)
    df["_name"] = df[name_col].astype(str).str.strip()
    df = df[(df["_upc"] != "") & (~df["_name"].isin(["", "nan", "None"]))]
    df = df.drop_duplicates(subset=["_upc"], keep="last")
    return dict(zip(df["_upc"], df["_name"])), latest


def load_barcode_to_sitegiant_name_map(file_id):
    """統整表 c -> sitegiant庫存SKU。空值不進 map；另回傳 SKU 空白的條碼集合。"""
    if not file_id:
        raise ValueError("找不到商品蝦皮麗嬰價格統整表")
    engine_kw = {"engine": "calamine"} if HAS_CALAMINE else {}
    df = pd.read_excel(download_gdrive_file_to_bytes(file_id), dtype=str, **engine_kw)
    df.columns = df.columns.astype(str).str.strip().str.replace("\n", "")
    if "c" not in df.columns or "sitegiant庫存SKU" not in df.columns:
        raise ValueError("統整表缺少「c」或「sitegiant庫存SKU」欄位")
    df["_bc"] = df["c"].map(clean_barcode)
    df["_name"] = df["sitegiant庫存SKU"].map(lambda x: str(x).strip() if pd.notna(x) else "")
    all_barcodes = set(df.loc[df["_bc"] != "", "_bc"])
    filled = df[(df["_bc"] != "") & (~df["_name"].isin(["", "nan", "None"]))]
    filled = filled.drop_duplicates(subset=["_bc"], keep="last")
    name_map = dict(zip(filled["_bc"], filled["_name"]))
    empty_sku_barcodes = all_barcodes - set(name_map.keys())
    return name_map, empty_sku_barcodes


def apply_sitegiant_name_from_summary(df, name_map, empty_sku_barcodes=None):
    """以國際條碼對 name_map，覆寫庫存貨品名稱。不改庫存SKU / 成本 / 稅款。"""
    empty_sku_barcodes = set(empty_sku_barcodes or [])
    out = df.copy()
    out.columns = out.columns.astype(str).str.strip().str.replace("\n", "")
    if "國際條碼" not in out.columns:
        raise ValueError("檔案缺少「國際條碼」欄位")
    if "庫存貨品名稱" not in out.columns:
        out["庫存貨品名稱"] = ""

    barcodes = out["國際條碼"].map(clean_barcode)
    mapped = barcodes.map(lambda x: name_map.get(x) if x else None)
    has_name = mapped.notna() & (mapped.astype(str).str.strip() != "")

    old_names = out["庫存貨品名稱"].where(out["庫存貨品名稱"].notna(), "").astype(str).str.strip()
    new_names = mapped.where(has_name, old_names).astype(str).str.strip()
    out.loc[has_name, "庫存貨品名稱"] = mapped[has_name]
    updated = int((has_name & (old_names != new_names)).sum())

    valid_bc = barcodes != ""
    unmatched = valid_bc & ~has_name
    is_empty_sku = unmatched & barcodes.isin(empty_sku_barcodes)
    is_not_found = unmatched & ~is_empty_sku
    return out, {
        "updated": updated,
        "not_found": barcodes[is_not_found].tolist(),
        "empty_sku": barcodes[is_empty_sku].tolist(),
    }


def build_price_summary_df(master_file_id, local_prod_id, shopee_master_id, sitegiant_batch_folder_id=None):
    """三表整合 + 以 c=UPC 填入 sitegiant庫存SKU（庫存貨品名稱）。"""
    engine_kw = {"engine": "calamine"} if HAS_CALAMINE else {}
    df_liying = pd.read_excel(download_gdrive_file_to_bytes(master_file_id), sheet_name="麗嬰國際產品總表", **engine_kw)
    df_p = pd.read_excel(download_gdrive_file_to_bytes(local_prod_id), sheet_name=0, **engine_kw)
    df_s = pd.read_excel(download_gdrive_file_to_bytes(shopee_master_id), sheet_name="蝦皮商品列表", **engine_kw)

    df_liying["條碼"] = df_liying["條碼"].astype(str).str.strip().str.split(".").str[0]
    df_p["自定義編碼"] = df_p["自定義編碼"].astype(str).str.strip().str.split(".").str[0]
    df_s["iSKU"] = df_s["iSKU"].astype(str).str.strip().str.split(".").str[0]

    if "商品名稱" in df_p.columns:
        df_p = df_p.rename(columns={"商品名稱": "內部商品名稱"})

    df_merge1 = pd.merge(df_p, df_s[["商品名稱", "iSKU", "GTIN", "價格"]], left_on="自定義編碼", right_on="iSKU", how="left")
    df_merge1 = df_merge1.rename(columns={"商品名稱": "蝦皮商品名稱", "GTIN": "蝦皮GTIN", "價格": "蝦皮售價"})
    df_merge1["c"] = df_merge1["c"].astype(str).str.strip().str.split(".").str[0]

    df_final = pd.merge(df_merge1, df_liying[["條碼", "零售價", "含稅"]], left_on="c", right_on="條碼", how="left")
    df_final = df_final.rename(columns={"零售價": "麗嬰零售價", "含稅": "麗嬰批發含稅價", "條碼": "麗嬰條碼"})
    df_final["麗嬰商品"] = df_final["麗嬰條碼"].apply(lambda x: None if pd.isna(x) else "v")

    for c in ["蝦皮售價", "麗嬰零售價", "麗嬰批發含稅價"]:
        df_final[c] = pd.to_numeric(df_final[c], errors="coerce")

    df_final["麗嬰零售八折"] = df_final["麗嬰零售價"] * 0.8
    df_final["麗嬰八折比蝦皮貴"] = df_final.apply(
        lambda r: "v" if (pd.notna(r["麗嬰零售八折"]) and pd.notna(r["蝦皮售價"]) and r["麗嬰零售八折"] > r["蝦皮售價"]) else None,
        axis=1,
    )
    df_final["麗嬰未稅價"] = df_final["麗嬰批發含稅價"].apply(lambda x: round(x / 1.05, 2) if pd.notna(x) else None)
    df_final["麗嬰稅款"] = df_final.apply(
        lambda r: round(r["麗嬰批發含稅價"] - r["麗嬰未稅價"], 2) if (pd.notna(r["麗嬰批發含稅價"]) and pd.notna(r["麗嬰未稅價"])) else None,
        axis=1,
    )

    name_map, batch_meta = {}, None
    if sitegiant_batch_folder_id:
        name_map, batch_meta = load_sitegiant_batch_name_map(sitegiant_batch_folder_id)
    df_final["sitegiant庫存SKU"] = df_final["c"].map(lambda x: name_map.get(clean_barcode(x)) if pd.notna(x) else None)

    return df_final.drop(columns=["iSKU"], errors="ignore"), batch_meta


def correct_shopee_isku(file_bytes):
    payload = _file_payload_bytes(file_bytes)
    df_shopee_raw = pd.read_excel(io.BytesIO(payload), header=None, engine="openpyxl")
    if df_shopee_raw.shape[1] >= 11:
        df_shopee_raw.drop(df_shopee_raw.columns[10], axis=1, inplace=True)
    shopee_headers = df_shopee_raw.iloc[2].astype(str).str.strip().tolist()
    df_shopee = df_shopee_raw.iloc[6:].copy()
    df_shopee.columns = shopee_headers
    df_shopee.reset_index(drop=True, inplace=True)

    def calc_isku_row(row):
        opt = str(row.get("商品選項貨號", "")).strip()
        main = str(row.get("主商品貨號", "")).strip()
        if opt in ["見選項", "null", "Null", "nan", "NaN", "None"]:
            opt = ""
        if main in ["見選項", "null", "Null", "nan", "NaN", "None"]:
            main = ""
        return opt if opt != "" else (main if main != "" else "蝦皮無iSKU")

    df_shopee["iSKU"] = df_shopee.apply(calc_isku_row, axis=1)
    df_shopee["original_index"] = df_shopee.index
    cols_list = list(df_shopee.columns)
    if "iSKU" in cols_list and "價格" in cols_list:
        cols_list.remove("iSKU")
        cols_list.insert(cols_list.index("價格"), "iSKU")
        df_shopee = df_shopee[cols_list]

    df_valid_isku = df_shopee[df_shopee["iSKU"] != "蝦皮無iSKU"].copy()
    df_isku_keep = df_valid_isku.sort_values(by=["iSKU", "價格", "original_index"]).drop_duplicates(subset=["iSKU"], keep="last")
    df_gtin_check = df_isku_keep.copy()
    df_gtin_check["GTIN_str"] = df_gtin_check["GTIN"].astype(str).str.strip().str.split(".").str[0]
    df_gtin_keep = df_gtin_check[~df_gtin_check["GTIN_str"].isin(["", "00", "0", "nan"])].sort_values(
        by=["GTIN_str", "價格", "original_index"]
    ).drop_duplicates(subset=["GTIN_str"], keep="last")
    return pd.concat(
        [df_gtin_keep, df_gtin_check[df_gtin_check["GTIN_str"].isin(["", "00", "0", "nan"])]]
    ).sort_values(by="original_index")


def taipei_now():
    return datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8)))


def shopee_imported_today(df_hist):
    if df_hist is None or df_hist.empty or "匯入時間" not in df_hist.columns:
        return False
    today = taipei_now().strftime("%Y-%m-%d")
    return df_hist["匯入時間"].astype(str).str.startswith(today).any()


def apply_shopee_isku_from_source(file_bytes, source_name, shopee_master_id, shopee_folder_id, shopee_file_name):
    payload = _file_payload_bytes(file_bytes)
    md5 = calculate_md5(payload)
    df_hist, _ = load_shopee_data(shopee_master_id)
    if df_hist is None:
        df_hist = pd.DataFrame(columns=["檔案名稱", "md5", "匯入時間"])
    if not df_hist.empty and "md5" in df_hist.columns and md5 in df_hist["md5"].astype(str).values:
        return {"ok": False, "reason": "duplicate", "md5": md5, "name": source_name, "df": None}

    df_clean = correct_shopee_isku(payload)
    imported_at = taipei_now().strftime("%Y-%m-%d %H:%M:%S")
    new_hist = pd.concat(
        [df_hist, pd.DataFrame([{"檔案名稱": source_name, "md5": md5, "匯入時間": imported_at}])],
        ignore_index=True,
    )
    saved = save_to_shopee_master_xlsm(
        {"蝦皮商品列表": df_clean, "匯入檔案": new_hist},
        shopee_master_id,
        shopee_folder_id,
        shopee_file_name,
    )
    return {
        "ok": bool(saved),
        "reason": None if saved else "save_failed",
        "md5": md5,
        "name": source_name,
        "imported_at": imported_at,
        "df": df_clean if saved else None,
    }