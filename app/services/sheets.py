"""
Google Sheets Service — IHK Upload
Mengelola koneksi dan operasi tulis ke Google Sheets via service account.
"""

import gspread
import time
import logging
import os
from collections import OrderedDict
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1
from gspread.exceptions import APIError
from app.main.brs_cols import BRS_CONFIG


_log = logging.getLogger(__name__)

# ─── Global Progress & Cache ──────────────────────────────────────────────────
UPLOAD_PROGRESS = {}  # { task_id: {"percent": 0~100, "message": "..."} }
_gspread_client = None

# Cache for document registry, portal links, and users
_CACHES = {
    'docs':  {'data': {}, 'ts': {}},  # {section: list}
    'klik':  {'data': None, 'ts': 0.0},
    'users': {'data': {}, 'ts': 0.0}
}

def set_progress(task_id: str, percent: int, message: str):
    if not task_id: return
    UPLOAD_PROGRESS[task_id] = {
        "percent": max(0, min(100, percent)),
        "message": message
    }

def get_progress(task_id: str) -> dict:
    return UPLOAD_PROGRESS.get(task_id, {"percent": 0, "message": "Menunggu respon server..."})

def clear_progress(task_id: str):
    UPLOAD_PROGRESS.pop(task_id, None)

# ─── Generic Cache Helper ─────────────────────────────────────────────────────

def _get_cached_data(cache_key: str, fetch_func, sub_key: str = None):
    """Generic helper to handle TTL caching for GSheets data."""
    from app.constants import DOCS_CACHE_TTL
    now = time.time()
    cache = _CACHES[cache_key]
    
    # Handle nested cache (like 'docs' per section)
    if sub_key is not None:
        if sub_key in cache['data'] and (now - cache['ts'].get(sub_key, 0)) < DOCS_CACHE_TTL:
            return cache['data'][sub_key]
        
        data = fetch_func()
        if data is not None:
            cache['data'][sub_key] = data
            cache['ts'][sub_key] = now
        return data or []
    
    # Handle single cache (like 'klik' or 'users')
    if cache['data'] is not None and (now - cache['ts']) < DOCS_CACHE_TTL:
        return cache['data']
        
    data = fetch_func()
    if data is not None:
        cache['data'] = data
        cache['ts'] = now
    return data or ({} if cache_key == 'users' else [])

# Spreadsheet IDs are now in app.constants

# ─── KUSTOMISASI IHK MAPPING ──────────────────────────────────────────────────
# Mapping dari BRS_CONFIG
SHEET_VALUE_MAP = {k: v for k, v in BRS_CONFIG['ihk']['required_cols'].items() if v in BRS_CONFIG['ihk']['sheets']}
FIXED_COLS = [v for k, v in BRS_CONFIG['ihk']['required_cols'].items() if v not in BRS_CONFIG['ihk']['sheets'] and k not in ('Tahun', 'Bulan')] + list(BRS_CONFIG['ihk']['optional_cols'].values())

CREDS_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'credentials.json')
SCOPES = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

# ─── Registry Dokumen ─────────────────────────────────────────────────────────

def get_docs(section: str, fallback: list) -> list:
    """Ambil daftar dokumen untuk section BRS dari Google Sheets registry."""
    from app.constants import DOCS_SPREADSHEET_ID, DOCS_SHEET_MAP

    if not DOCS_SPREADSHEET_ID: return fallback
    sheet_name = DOCS_SHEET_MAP.get(section)
    if not sheet_name: return fallback

    def _fetch():
        try:
            spreadsheet = _get_spreadsheet(DOCS_SPREADSHEET_ID)
            worksheet = spreadsheet.worksheet(sheet_name)
            rows = worksheet.get_all_values()
            if not rows or len(rows) < 2: return None
            
            header = [h.strip().lower() for h in rows[0]]
            
            def _idx(names, default):
                for n in names:
                    if n in header: return header.index(n)
                return default

            idx_label = _idx(['nama dokumen', 'label', 'nama'], 0)
            idx_kat   = _idx(['kategori', 'category', 'group'], 1)
            idx_desc  = _idx(['deskripsi', 'description', 'desc'], 2)
            idx_url   = _idx(['link', 'url'], 3)

            grouped = OrderedDict()
            kat_icons = {'master': 'bi-pencil-fill', 'brs': 'bi-file-earmark-text-fill', 'folder': 'bi-folder2-open'}
            item_icons = {'master': 'bi-file-earmark-excel-fill', 'brs': 'bi-file-earmark-text', 'folder': 'bi-folder'}

            for row in rows[1:]:
                if not any(row): continue
                label = row[idx_label].strip() if len(row) > idx_label else ''
                if not label: continue
                
                kat = row[idx_kat].strip() if len(row) > idx_kat else 'Lainnya'
                desc = row[idx_desc].strip() if len(row) > idx_desc else ''
                url = row[idx_url].strip() if len(row) > idx_url else ''
                
                if kat not in grouped: grouped[kat] = []
                grouped[kat].append({
                    'label': label, 'desc': desc, 'url': url,
                    'icon': item_icons.get(kat.lower(), 'bi-box-arrow-up-right')
                })

            return [
                {'kategori': k, 'icon': kat_icons.get(k.lower(), 'bi-collection'), 'dokumen': d}
                for k, d in grouped.items()
            ]
        except Exception as e:
            _log.error(f'get_docs({section}) error: {e}')
            return None

    return _get_cached_data('docs', _fetch, sub_key=section) or fallback

def clear_docs_cache(section: str = None):
    """Paksa refresh cache registry dokumen."""
    if section:
        _CACHES['docs']['data'].pop(section, None)
        _CACHES['docs']['ts'].pop(section, None)
    else:
        _CACHES['docs']['data'].clear()
        _CACHES['docs']['ts'].clear()


# ─── Klik Distribusi ──────────────────────────────────────────────────────────

def get_klik_links() -> list:
    """Fetch daftar link portal dari Google Sheets."""

    if not KLIK_SPREADSHEET_ID: return []

    def _fetch():
        try:
            spreadsheet = _get_spreadsheet(KLIK_SPREADSHEET_ID)
            worksheet = spreadsheet.worksheet('Klik')
            rows = worksheet.get_all_values()
            if not rows or len(rows) < 2: return None
            
            header = [h.strip().lower() for h in rows[0]]
            def _idx(names, default):
                for n in names:
                    if n in header: return header.index(n)
                return default

            idx_kat  = _idx(['kategori', 'group'], 0)
            idx_nama = _idx(['nama', 'label', 'title'], 1)
            idx_link = _idx(['link', 'url'], 2)
            idx_jns  = _idx(['jenis', 'type'], 3)
            idx_ket  = _idx(['keterangan', 'desc', 'description'], 5)
            idx_thn  = _idx(['tahun', 'year', 'thn', 'periode'], 6)

            grouped = OrderedDict()
            jenis_icons = {
                'internal app': 'bi-shield-lock', 'dashboard': 'bi-speedometer2',
                'drive bps': 'bi-cloud-fill', 'g-drive': 'bi-folder2-open',
                'g-sheets': 'bi-grid-3x3-gap-fill', 'web app': 'bi-globe',
                'web entry': 'bi-keyboard', 'website': 'bi-window',
            }
            kat_colors = {
                'Utama': 'primary', 'Monitoring': 'success', 'Statistik Harga': 'danger',
                'Statistik Distribusi dan Jasa': 'purple', 'Lainnya': 'secondary',
            }

            for row in rows[1:]:
                if not any(row): continue
                nama = row[idx_nama].strip() if len(row) > idx_nama else ''
                if not nama: continue

                kat = row[idx_kat].strip() if len(row) > idx_kat else 'Lainnya'
                jenis = row[idx_jns].strip() if len(row) > idx_jns else ''
                tahun_raw = row[idx_thn].strip() if len(row) > idx_thn else ''
                
                try:
                    tahun = int(float(tahun_raw)) if tahun_raw else None
                except (ValueError, TypeError):
                    tahun = None

                if kat not in grouped: grouped[kat] = []
                grouped[kat].append({
                    'nama': nama,
                    'link': row[idx_link].strip() if len(row) > idx_link else '#',
                    'jenis': jenis,
                    'icon': jenis_icons.get(jenis.lower(), 'bi-box-arrow-up-right'),
                    'keterangan': row[idx_ket].strip() if len(row) > idx_ket else '',
                    'tahun': tahun,
                })

            return [
                {'kategori': k, 'color': kat_colors.get(k, 'secondary'), 'links': d}
                for k, d in grouped.items()
            ]
        except Exception as e:
            _log.error(f'get_klik_links() error: {e}')
            return None

def clear_klik_cache():
    """Paksa refresh cache Klik Distribusi."""
    _CACHES['klik']['data'] = None
    _CACHES['klik']['ts'] = 0.0


# ─── Data User ────────────────────────────────────────────────────────────────

def get_users() -> dict:
    """Fetch daftar user dan hak akses dari Google Sheets."""
    from app.constants import USERS_SPREADSHEET_ID, USERS_SHEET_NAME
    if not USERS_SPREADSHEET_ID: return {}

    def _fetch():
        try:
            spreadsheet = _get_spreadsheet(USERS_SPREADSHEET_ID)
            worksheet = spreadsheet.worksheet(USERS_SHEET_NAME)
            rows = worksheet.get_all_values()
            if not rows or len(rows) < 2: return None

            header = [str(h).strip().lower() for h in rows[0]]
            def _idx(names):
                for n in names:
                    if n in header: return header.index(n)
                return -1

            idx_user  = _idx(['username', 'user'])
            idx_nama  = _idx(['nama', 'name'])
            idx_pass  = _idx(['password', 'sandi'])
            idx_ihk   = _idx(['akses ihk', 'ihk'])
            idx_exim  = _idx(['akses ekspor impor', 'akses exim', 'ekspor impor', 'exim'])
            idx_ntp   = _idx(['akses ntp', 'ntp'])
            idx_trans = _idx(['akses transportasi', 'transportasi'])
            idx_pari  = _idx(['akses pariwisata', 'pariwisata'])

            if idx_user == -1 or idx_pass == -1:
                _log.error("Kolom 'Username' atau 'Password' tidak ditemukan.")
                return None

            users_dict = {}
            for row in rows[1:]:
                if not any(row): continue
                username_raw = row[idx_user].strip()
                password = row[idx_pass].strip()
                if not username_raw or not password: continue
                
                username = username_raw.lower()
                def _bool(idx):
                    return str(row[idx]).strip().upper() == 'TRUE' if idx != -1 and len(row) > idx else False

                users_dict[username] = {
                    'nama': row[idx_nama].strip() if idx_nama != -1 and len(row) > idx_nama else username_raw,
                    'password': password,
                    'akses_ihk': _bool(idx_ihk),
                    'akses_ekspor_impor': _bool(idx_exim),
                    'akses_ntp': _bool(idx_ntp),
                    'akses_transportasi': _bool(idx_trans),
                    'akses_pariwisata': _bool(idx_pari),
                }
            return users_dict
        except Exception as e:
            _log.error(f'get_users() error: {e}')
            return None

    return _get_cached_data('users', _fetch)

def clear_users_cache():
    """Paksa refresh cache user."""
    _CACHES['users']['data'] = {}
    _CACHES['users']['ts'] = 0.0


# ─── Koneksi ──────────────────────────────────────────────────────────────────

_gspread_client = None

def get_client() -> gspread.Client:
    """
    Buat atau kembalikan koneksi gspread menggunakan service account.
    Caches the client to avoid repeated authorizations.
    """
    global _gspread_client
    if _gspread_client:
        return _gspread_client

    creds_json = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    if creds_json:
        import json
        info = json.loads(creds_json)
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file(CREDS_PATH, scopes=SCOPES)
    
    _gspread_client = gspread.authorize(creds)
    return _gspread_client

def _get_spreadsheet(spreadsheet_id: str) -> gspread.spreadsheet.Spreadsheet:
    """Helper method to get spreadsheet by ID using the service account client."""
    client = get_client()
    try:
        return client.open_by_key(spreadsheet_id)
    except APIError as e:
        if '429' in str(e):
            _log.warning("Quota exceeded saat membuka spreadsheet. Menunggu 10 detik...")
            time.sleep(10)
            return client.open_by_key(spreadsheet_id)
        raise e


# ─── Helper ───────────────────────────────────────────────────────────────────

def _get_master_hs_data(spreadsheet: gspread.spreadsheet.Spreadsheet) -> dict:
    """
    Mengambil mapping dari sheet 'MASTER' untuk referensi HS Code ke OIL.
    Map format: {'HS_CODE_STRING': 'OIL_VALUE'}
    """
    try:
        worksheet = spreadsheet.worksheet('MASTER')
        rows = worksheet.get_all_values()
        if not rows:
            return {}
            
        header = rows[0]
        try:
            hs_idx = header.index('HS Code BTKI_2017')
            oil_idx = header.index('OIL')
        except ValueError as e:
            _log.error(f"Kolom wajib di sheet MASTER tidak ditemukan: {e}")
            return {}
            
        mapping = {}
        for row in rows[1:]:
            if len(row) > max(hs_idx, oil_idx):
                hs_code = str(row[hs_idx]).strip()
                if hs_code:
                    mapping[hs_code] = str(row[oil_idx]).strip()
        return mapping
    except Exception as e:
        _log.error(f"Gagal mendapat data MASTER: {e}")
        return {}


def _col_index(header_row: list, name: str) -> int | None:
    """Kembalikan 0-based index kolom dari header_row, atau None jika tidak ada."""
    try:
        return header_row.index(name)
    except ValueError:
        return None


def _retry_on_429(func):
    """Decorator internal untuk handle quota exceeded (429)."""
    def wrapper(*args, **kwargs):
        task_id = kwargs.pop('_task_id', None)
        max_retries = 5
        delay = 10
        for i in range(max_retries):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                err_msg = str(e)
                if '429' in err_msg and i < max_retries - 1:
                    _log.warning(f"Quota exceeded (429). Menunggu {delay} detik (retry {i+1}/{max_retries})...")
                    if task_id:
                        curr = get_progress(task_id).get('percent', 0)
                        set_progress(task_id, curr, f"Menunggu {delay} detik karena limit Google API (Retry {i+1}/{max_retries})")
                    time.sleep(delay)
                    delay += 10
                    continue
                raise e
        return func(*args, **kwargs)
    return wrapper


def _to_number(val):
    """
    Konversi nilai string ke float untuk disimpan sebagai angka di GSheets.
    Mendukung format Indonesia (koma sebagai desimal, misal '102,26').
    Jika tidak bisa dikonversi, kembalikan nilai aslinya.
    """
    if val is None or str(val).strip() == '':
        return ''
    try:
        clean = str(val).strip().replace(',', '.')
        result = float(clean)
        # Kembalikan int jika bilangan bulat, float jika desimal
        return int(result) if result == int(result) else result
    except (ValueError, TypeError):
        return val


def find_or_create_col(worksheet: gspread.Worksheet, header_row: list[str], col_name: str) -> tuple[int, bool]:
    """
    Cari kolom `col_name` di header_row.
    Jika tidak ada, kembalikan index baru dan flag True (perlu update).
    """
    if col_name in header_row:
        return header_row.index(col_name) + 1, False

    new_col_1based = len(header_row) + 1
    return new_col_1based, True


# ─── Per-sheet Logic ──────────────────────────────────────────────────────────

def _build_row_lookup(data_rows: list[list[str]], col_indices: dict[str, int]) -> dict[tuple[str, str], int]:
    """Membangun tabel pencarian untuk menemukan baris yang sudah ada berdasarkan Kode Kota dan Kode Komoditas."""
    lookup_table = {}
    for i, row in enumerate(data_rows, start=5):
        kode_kota = row[col_indices['Kd.Kota']] if col_indices['Kd.Kota'] is not None and col_indices['Kd.Kota'] < len(row) else ''
        kode_komoditas = row[col_indices['Kode']] if col_indices['Kode'] is not None and col_indices['Kode'] < len(row) else ''
        if kode_kota or kode_komoditas:
            lookup_table[(str(kode_kota).strip(), str(kode_komoditas).strip())] = i
    return lookup_table

def _create_ihk_row(row, col_indices, row_width, val_col_idx, cell_value):
    """Helper to build a single IHK data row based on BRS_CONFIG mapping."""
    row_data = [''] * row_width
    kd_cola = BRS_CONFIG['ihk']['keys'][0]
    kd_colb = BRS_CONFIG['ihk']['keys'][1]
    
    def _set(col_key, val):
        idx = col_indices.get(col_key)
        if idx is not None and idx < row_width:
            row_data[idx] = val

    # Set required columns
    _set(BRS_CONFIG['ihk']['required_cols'][kd_cola], str(row.get(kd_cola, '')).strip())
    _set(BRS_CONFIG['ihk']['required_cols'][kd_colb], str(row.get(kd_colb, '')).strip())
    
    # Set other fixed and optional columns
    for ex_col, sh_col in BRS_CONFIG['ihk']['required_cols'].items():
        if sh_col in FIXED_COLS and ex_col not in BRS_CONFIG['ihk']['keys'] and ex_col not in ('Tahun', 'Bulan'):
            _set(sh_col, str(row.get(ex_col, '')).strip())
            
    for ex_col, sh_col in BRS_CONFIG['ihk']['optional_cols'].items():
        _set(sh_col, str(row.get(ex_col, '')).strip())
    
    # Set the actual value for the current period
    if val_col_idx < row_width:
        row_data[val_col_idx] = cell_value
        
    return row_data

def _upsert_ihk_sheet(worksheet: gspread.Worksheet, dataframe, col_name: str, excel_col_name: str) -> tuple[int, int]:
    """Insert/update data dari DataFrame ke worksheet IHK."""
    all_values = worksheet.get_all_values()
    header_row = all_values[2] if len(all_values) > 2 else []
    col_indices = {name: _col_index(header_row, name) for name in FIXED_COLS}

    # 1. Ensure period column exists
    val_col_1based, is_new = find_or_create_col(worksheet, header_row, col_name)
    val_col_idx = val_col_1based - 1
    
    if is_new:
        row4 = all_values[3] if len(all_values) > 3 else []
        num_filled = sum(1 for x in row4 if str(x).strip())
        worksheet.batch_update([
            {'range': rowcol_to_a1(3, val_col_1based), 'values': [[col_name]]},
            {'range': rowcol_to_a1(4, val_col_1based), 'values': [[num_filled + 1]]},
        ])
        while len(header_row) < val_col_1based: header_row.append('')
        header_row[val_col_idx] = col_name

    # 2. Build lookup and process data
    data_rows = all_values[4:] 
    lookup = _build_row_lookup(data_rows, col_indices)
    
    batch_updates, new_rows = [], []
    inserted_count = updated_count = 0
    max_col_0 = max((v for v in col_indices.values() if v is not None), default=0)
    row_width = max(max_col_0, val_col_idx) + 1

    kd_cola, kd_colb = BRS_CONFIG['ihk']['keys']
    
    for _, row in dataframe.iterrows():
        kd_kota = str(row.get(kd_cola, '')).strip()
        kd_komo = str(row.get(kd_colb, '')).strip()
        cell_value = _to_number(row.get(excel_col_name, ''))
        key = (kd_kota, kd_komo)

        if key in lookup:
            batch_updates.append({
                'range': rowcol_to_a1(lookup[key], val_col_1based),
                'values': [[cell_value]],
            })
            updated_count += 1
        else:
            new_row = _create_ihk_row(row, col_indices, row_width, val_col_idx, cell_value)
            new_rows.append(new_row)
            lookup[key] = 5 + len(data_rows) + len(new_rows) - 1
            inserted_count += 1

    # 3. Execute updates
    if batch_updates:
        worksheet.batch_update(batch_updates)

    if new_rows:
        next_row = max(len(all_values) + 1, 5)
        last_row = next_row + len(new_rows) - 1
        if last_row > worksheet.row_count:
            worksheet.add_rows(last_row - worksheet.row_count + 100)
            
        range_name = f"{rowcol_to_a1(next_row, 1)}:{rowcol_to_a1(last_row, len(new_rows[0]))}"
        worksheet.update(new_rows, range_name)

    return inserted_count, updated_count

# ─── IHK Orchestrator ─────────────────────────────────────────────────────────────

def process_ihk_upload(dataframe, col_name: str, task_id: str = None) -> dict:
    """Upload IHK DataFrame ke 7 sheet Google Sheets sekaligus."""
    from app.constants import IHK_SPREADSHEET_ID
    set_progress(task_id, 10, "Menyambungkan ke Google Sheets API...")
    spreadsheet = _get_spreadsheet(IHK_SPREADSHEET_ID)

    total_inserted = total_updated = 0
    details = []
    num_sheets = len(SHEET_VALUE_MAP)
    set_progress(task_id, 20, f"Mempersiapkan {num_sheets} sheet...")

    for idx, (ex_col, sh_name) in enumerate(SHEET_VALUE_MAP.items(), 1):
        pct = 20 + int(80 * (idx - 1) / num_sheets)
        set_progress(task_id, pct, f"Memproses sheet '{sh_name}' ({idx}/{num_sheets})...")
        
        worksheet = spreadsheet.worksheet(sh_name)
        # Handle quota limits with retry
        func = _retry_on_429(_upsert_ihk_sheet)
        ins, upd = func(worksheet, dataframe, col_name, ex_col, _task_id=task_id)
        
        total_inserted += ins
        total_updated  += upd
        details.append({'sheet': sh_name, 'inserted': ins, 'updated': upd})
        time.sleep(1) # Safety delay

    set_progress(task_id, 100, "Selesai!")
    return {'inserted': total_inserted, 'updated': total_updated, 'details': details}

# ─── Ekspor Upload ────────────────────────────────────────────────────────────

def _ensure_ekspor_yymm_column(worksheet: gspread.Worksheet, header_row: list[str], yymm_label: str, batch_header_updates: list) -> int:
    """Pastikan kolom YYMM ada di sheet. Jika tidak ada, tambahkan ke header_row dan batch_header_updates."""
    if not header_row:
        return -1
        
    # Check using string comparison to be type-agnostic (handles both existing int/str in sheet)
    str_header = [str(h).strip() for h in header_row]
    if yymm_label in str_header:
        return str_header.index(yymm_label) + 1
        
    yymm_idx = len(header_row) + 1
    
    # Perbaikan: Pastikan grid sheet cukup luas sebelum batch_update
    if yymm_idx > worksheet.col_count:
        worksheet.add_cols(max(1, yymm_idx - worksheet.col_count))
        
    try:
        yymm_val = int(yymm_label)
    except:
        yymm_val = yymm_label

    batch_header_updates.append({
        'range': rowcol_to_a1(1, yymm_idx),
        'values': [[yymm_val]]
    })
    header_row.append(str(yymm_val)) # Keep as string in local header_row for consistent lookup
    return yymm_idx

def _upsert_recap_sheet(worksheet: gspread.Worksheet, months_recap_data: dict) -> tuple[int, int]:
    """
    Update atau insert data agregat ke sheet RECAP.
    months_recap_data: { 'yymm': {'MIGAS': val, 'NON MIGAS': val, 'Total': val} }
    """
    all_values = worksheet.get_all_values()
    if not all_values:
        header = ['Periode', 'Migas', 'Non Migas', 'Total']
        worksheet.append_row(header)
        all_values = [header]
    
    header_row = all_values[0]
    # Pastikan header sesuai
    cols = ['Periode', 'Migas', 'Non Migas', 'Total']
    col_map = {}
    for c in cols:
        if c not in header_row:
            idx = len(header_row) + 1
            worksheet.update_cell(1, idx, c)
            header_row.append(c)
        col_map[c] = header_row.index(c) + 1

    lookup = {}
    for i, row in enumerate(all_values[1:], start=2):
        if row and len(row) > 0:
            lookup[str(row[0]).strip()] = i
            
    batch_updates = []
    new_rows = []
    
    for yymm, vals in months_recap_data.items():
        # Force string agar di Google Sheets dibaca sebagai text (menggunakan tanda petik tunggal)
        yymm_val = f"{yymm}"
            
        if yymm in lookup:
            row_idx = lookup[yymm]
            batch_updates.append({'range': rowcol_to_a1(row_idx, col_map['Migas']), 'values': [[float(vals.get('MIGAS', 0))]]})
            batch_updates.append({'range': rowcol_to_a1(row_idx, col_map['Non Migas']), 'values': [[float(vals.get('NON MIGAS', 0))]]})
            batch_updates.append({'range': rowcol_to_a1(row_idx, col_map['Total']), 'values': [[float(vals.get('Total', 0))]]})
        else:
            row_data = [''] * len(header_row)
            row_data[col_map['Periode']-1] = yymm_val
            row_data[col_map['Migas']-1] = float(vals.get('MIGAS', 0))
            row_data[col_map['Non Migas']-1] = float(vals.get('NON MIGAS', 0))
            row_data[col_map['Total']-1] = float(vals.get('Total', 0))
            new_rows.append(row_data)
            
    if batch_updates:
        worksheet.batch_update(batch_updates)
    if new_rows:
        worksheet.append_rows(new_rows)
        
    # Sort by Periode (Column 1)
    try:
        # Use batch_update for sorting rows reliably
        final_row_count = worksheet.row_count
        if final_row_count > 1:
            body = {
                "requests": [{
                    "sortRange": {
                        "range": {
                            "sheetId": worksheet.id,
                            "startRowIndex": 1,
                            "endRowIndex": final_row_count,
                            "startColumnIndex": 0,
                            "endColumnIndex": 4
                        },
                        "sortSpecs": [
                            {"dimensionIndex": 0, "sortOrder": "ASCENDING"}
                        ]
                    }
                }]
            }
            worksheet.spreadsheet.batch_update(body)
    except Exception as e:
        _log.error(f"Gagal mengurutkan sheet RECAP: {e}")
        
    return len(new_rows), len(months_recap_data) - len(new_rows)

def _migrate_exim_header(worksheet, header_row, sheet_name, upload_type):
    """Migrate semantic keys in header to real DBF column names."""
    all_dbf = {}
    for mod in ('ekspor', 'impor'):
        all_dbf.update(BRS_CONFIG[mod]['required_dbf_cols'])
        
    semantic_cols = [(i, h) for i, h in enumerate(header_row) if h in all_dbf]
    if not semantic_cols: return header_row

    updates = []
    for idx, semantic in semantic_cols:
        real = all_dbf[semantic]
        _log.info(f'Migrasi header {semantic!r} -> {real!r} di sheet {sheet_name!r}')
        header_row[idx] = real
        updates.append({'range': rowcol_to_a1(1, idx + 1), 'values': [[real]]})
    
    if updates: worksheet.batch_update(updates)
    
    # Remove accidental duplicates from failed previous runs
    seen, to_delete = {}, []
    for i, h in enumerate(header_row):
        if h in seen: to_delete.append(i)
        else: seen[h] = i
        
    for i in sorted(to_delete, reverse=True):
        _log.info(f"Hapus kolom duplikat '{header_row[i]}' di {sheet_name}")
        worksheet.delete_columns(i + 1)
        
    return worksheet.get_all_values()[0] if to_delete else header_row

def _sort_exim_sheet(worksheet, key_indices, header_row):
    """Sort rows by keys and columns by YYMM labels."""
    try:
        # 1. Row sorting
        if worksheet.row_count > 1:
            worksheet.spreadsheet.batch_update({
                "requests": [{
                    "sortRange": {
                        "range": {"sheetId": worksheet.id, "startRowIndex": 1, "endRowIndex": worksheet.row_count, "startColumnIndex": 0, "endColumnIndex": len(header_row)},
                        "sortSpecs": [{"dimensionIndex": idx, "sortOrder": "ASCENDING"} for idx in key_indices]
                    }
                }]
            })
        
        # 2. Column sorting (YYMM horizontal)
        curr_header = worksheet.get_all_values()[0]
        periods = sorted([(v, i) for i, v in enumerate(curr_header) if str(v).isdigit() and len(str(v)) == 4], key=lambda x: x[0])
        if not periods: return

        start_offset = min(p[1] for p in periods)
        moves = []
        sim_header = list(curr_header)
        for target_pos, (label, _) in enumerate(periods):
            target_idx = start_offset + target_pos
            actual_idx = sim_header.index(label)
            if actual_idx != target_idx:
                moves.append({
                    "moveDimension": {
                        "source": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": actual_idx, "endIndex": actual_idx + 1},
                        "destinationIndex": target_idx if target_idx < actual_idx else target_idx + 1
                    }
                })
                sim_header.insert(target_idx, sim_header.pop(actual_idx))
        
        if moves: worksheet.spreadsheet.batch_update({"requests": moves})
    except Exception as e:
        _log.error(f"Gagal sorting: {e}")

def _upsert_exim_sheet_bulk(worksheet: gspread.Worksheet, months_data: dict, key_col_name: str | list[str], val_col: str, sheet_name: str = None, upload_type: str = 'ekspor') -> tuple[int, int]:
    """Bulk upsert aggregated data to Export/Import sheets."""
    all_values = worksheet.get_all_values()
    header_row = all_values[0] if all_values else []
    key_names = [key_col_name] if isinstance(key_col_name, str) else key_col_name
    prov_col = 'KodeProvTujuan' if upload_type == 'impor' else 'KodeProvAsal'

    # 1. Initialize/Migrate Header
    if not header_row:
        sample = list(months_data.keys())[0] if months_data else 'YYMM'
        header_row = key_names + ([prov_col] if sheet_name == 'Pelabuhan' else []) + [sample]
        worksheet.append_row(header_row)
    
    header_row = _migrate_exim_header(worksheet, header_row, sheet_name, upload_type)
    all_values = worksheet.get_all_values()
    
    # 2. Setup indices
    prov_idx = None
    if sheet_name == 'Pelabuhan':
        if prov_col not in header_row:
            primary_idx = header_row.index(key_names[0])
            _log.info(f"Menyisipkan kolom {prov_col}...")
            worksheet.insert_cols([[prov_col]], primary_idx + 2)
            all_values = worksheet.get_all_values()
            header_row = all_values[0]
        prov_idx = header_row.index(prov_col)
    
    key_indices = [header_row.index(k) for k in key_names]

    # 3. Process Data
    lookup = {tuple(str(row[i]).strip() for i in key_indices): i + 1 for i, row in enumerate(all_values[1:]) if any(row)}
    batch_updates, header_updates, new_rows_map = [], [], {}
    inserted = updated = 0

    for yymm, df_group in months_data.items():
        yymm_idx = _ensure_ekspor_yymm_column(worksheet, header_row, yymm, header_updates)
        width = max(len(header_row), max(key_indices) + 1, (prov_idx + 1) if prov_idx is not None else 0)

        for _, row in df_group.iterrows():
            key = tuple(str(row.get(k, '')).strip() for k in key_names)
            val = float(row[val_col])

            if key in lookup:
                row_idx = lookup[key]
                batch_updates.append({'range': rowcol_to_a1(row_idx, yymm_idx), 'values': [[val]]})
                if prov_idx is not None:
                    curr_row = all_values[row_idx-1] if row_idx-1 < len(all_values) else []
                    if prov_idx >= len(curr_row) or not str(curr_row[prov_idx]).strip():
                        batch_updates.append({'range': rowcol_to_a1(row_idx, prov_idx + 1), 'values': [[key[0][:2]]]})
                updated += 1
            else:
                if key not in new_rows_map:
                    rd = [''] * width
                    for i, idx in enumerate(key_indices): rd[idx] = key[i]
                    if prov_idx is not None: rd[prov_idx] = key[0][:2]
                    new_rows_map[key] = rd
                    inserted += 1
                
                if yymm_idx > len(new_rows_map[key]):
                    new_rows_map[key].extend([''] * (yymm_idx - len(new_rows_map[key])))
                new_rows_map[key][yymm_idx - 1] = val

    # 4. Execute updates
    if header_updates: worksheet.batch_update(header_updates)
    if batch_updates: worksheet.batch_update(batch_updates)
    if new_rows_map:
        new_rows = list(new_rows_map.values())
        max_c = max(len(r) for r in new_rows)
        for r in new_rows: r.extend([''] * (max_c - len(r)))
        
        start_r = len(all_values) + 1
        last_r = start_r + len(new_rows) - 1
        if last_r > worksheet.row_count: worksheet.add_rows(last_r - worksheet.row_count + 50)
        worksheet.update(new_rows, f"{rowcol_to_a1(start_r, 1)}:{rowcol_to_a1(last_r, max_c)}")

    _sort_exim_sheet(worksheet, key_indices, header_row)
    return inserted, updated

def _process_exim_orchestrator(dataframe, upload_type: str, task_id: str = None) -> dict:
    """Internal shared orchestrator for Ekspor and Impor uploads."""
    import pandas as pd
    from app.constants import EKSPOR_IMPOR_SPREADSHEET_ID, IMPOR_SPREADSHEET_ID
    
    config = BRS_CONFIG[upload_type]
    ss_id = EKSPOR_IMPOR_SPREADSHEET_ID if upload_type == 'ekspor' else IMPOR_SPREADSHEET_ID
    
    set_progress(task_id, 10, "Menyambungkan ke Google Sheets API...")
    spreadsheet = _get_spreadsheet(ss_id)
    
    _dbf = config['required_dbf_cols']
    val_col = _dbf['KeyNilai']
    dataframe['KeyNilai'] = pd.to_numeric(dataframe['KeyNilai'], errors='coerce').fillna(0)
    
    # 1. Prepare month groups
    set_progress(task_id, 20, "Mempersiapkan pengelompokan baris per bulan...")
    months_payload = {}
    for (year, month), group in dataframe.groupby(['KeyTahun', 'KeyBulan']):
        try:
            label = str(int(float(year)))[-2:] + str(int(float(month))).zfill(2)
            months_payload[label] = group
        except: continue

    if not months_payload:
        return {'inserted': 0, 'updated': 0, 'details': [{'error': 'No valid month/year data found'}]}

    total_inserted = total_updated = 0
    details = []
    targets = list(config['targets'].items())
    
    # 2. Process data sheets
    for idx, (sh_name, group_key) in enumerate(targets, 1):
        set_progress(task_id, 30 + (idx-1)*20, f"Memproses sheet '{sh_name}'...")
        
        keys_list = [group_key] if isinstance(group_key, str) else group_key
        rename_map = {k: _dbf[k] for k in keys_list}
        rename_map['KeyNilai'] = val_col
        
        prepared_months = {}
        for yymm, group in months_payload.items():
            summed = group.groupby(keys_list)['KeyNilai'].sum().reset_index()
            prepared_months[yymm] = summed.rename(columns=rename_map)
        
        real_key_col = [_dbf[k] for k in keys_list]
        if len(real_key_col) == 1: real_key_col = real_key_col[0]
        
        worksheet = spreadsheet.worksheet(sh_name)
        ins, upd = _retry_on_429(_upsert_exim_sheet_bulk)(worksheet, prepared_months, real_key_col, val_col, sh_name, upload_type=upload_type)
        total_inserted += ins
        total_updated += upd
        details.append({'sheet': sh_name, 'inserted': ins, 'updated': upd})
        time.sleep(1)

    # 3. Process RECAP sheet
    set_progress(task_id, 90, "Memproses sheet RECAP...")
    try:
        hs_map = _get_master_hs_data(spreadsheet)
        df_recap = dataframe.copy()
        df_recap['OIL'] = df_recap['KeyKodeHS'].map(lambda x: hs_map.get(str(x).strip(), 'NON MIGAS'))
        df_recap['OIL'] = df_recap['OIL'].apply(lambda x: 'MIGAS' if 'MIGAS' in str(x).upper() and 'NON' not in str(x).upper() else 'NON MIGAS')
        
        recap_payload = {}
        for (year, month), group in df_recap.groupby(['KeyTahun', 'KeyBulan']):
            try:
                yymm = str(int(float(year)))[-2:] + str(int(float(month))).zfill(2)
                oil_sums = group.groupby('OIL')['KeyNilai'].sum().to_dict()
                recap_payload[yymm] = {
                    'MIGAS': oil_sums.get('MIGAS', 0),
                    'NON MIGAS': oil_sums.get('NON MIGAS', 0),
                    'Total': group['KeyNilai'].sum()
                }
            except: continue
            
        if recap_payload:
            ins_r, upd_r = _retry_on_429(_upsert_recap_sheet)(spreadsheet.worksheet('RECAP'), recap_payload)
            total_inserted += ins_r
            total_updated += upd_r
            details.append({'sheet': 'RECAP', 'inserted': ins_r, 'updated': upd_r})
    except Exception as e:
        _log.error(f"Gagal update RECAP: {e}")
        details.append({'sheet': 'RECAP', 'error': str(e)})

    set_progress(task_id, 100, "Selesai!")
    return {'inserted': total_inserted, 'updated': total_updated, 'details': details}

def process_ekspor_upload(dataframe, task_id: str = None) -> dict:
    """Menerima dataframe ekspor dan menguploadnya."""
    return _process_exim_orchestrator(dataframe, 'ekspor', task_id)

def process_impor_upload(dataframe, task_id: str = None) -> dict:
    """Menerima dataframe impor dan menguploadnya."""
    return _process_exim_orchestrator(dataframe, 'impor', task_id)
