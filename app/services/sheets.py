"""
Google Sheets Service — IHK Upload
Mengelola koneksi dan operasi tulis ke Google Sheets via service account.
"""

import os
import logging

_log = logging.getLogger(__name__)
import gspread
import time
from gspread.utils import rowcol_to_a1
from gspread.exceptions import APIError
from google.oauth2.service_account import Credentials

# ─── Global Progress Tracking ─────────────────────────────────────────────────
# Menyimpan status progress upload Excel/DBF secara kasat-mata ke frontend via AJAX
# Format: { "task_id": {"percent": 0~100, "message": "Memproses..."} }
UPLOAD_PROGRESS = {}

def set_progress(task_id: str, percent: int, message: str):
    if not task_id:
        return
    # Pastikan nilai 0-100
    percent = max(0, min(100, percent))
    UPLOAD_PROGRESS[task_id] = {
        "percent": percent,
        "message": message
    }

def get_progress(task_id: str) -> dict:
    if not task_id or task_id not in UPLOAD_PROGRESS:
        return {"percent": 0, "message": "Menunggu respon server..."}
    return UPLOAD_PROGRESS[task_id]

def clear_progress(task_id: str):
    if task_id in UPLOAD_PROGRESS:
        del UPLOAD_PROGRESS[task_id]

# ─── Konstanta ────────────────────────────────────────────────────────────────

IHK_SPREADSHEET_ID = '1iB0DMQw7kjVzl5PgrVsb1wtQIlRF3QL-pZHYmQ0JySQ'
EKSPOR_IMPOR_SPREADSHEET_ID = '15FRr5c0HuhED1DpyMwfy6xXCxDXDXGMSiMYrk3a0ui0'
IMPOR_SPREADSHEET_ID = '1a3h0w4zE_kdvszM15H_UPWQ76QykSADIGOY-zXyTqzc'

# ─── KUSTOMISASI MAPPING ──────────────────────────────────────────────────────
#
# Format: 'Nama kolom di file Excel' : 'Nama sheet di Google Sheets'
#
# Sesuaikan bagian kanan (nama sheet) dengan nama sheet ASLI di spreadsheet kamu.
# Nama sheet bersifat case-sensitive!
#
SHEET_VALUE_MAP = {
    'IHK':        'IHK',          # kolom Excel "IHK"        → sheet "IHK"
    'Inflasi MtM': 'MTM',         # kolom Excel "Inflasi MtM" → sheet "MTM"
    'Inflasi YtD': 'YTD',         # kolom Excel "Inflasi YtD" → sheet "YTD"
    'Inflasi YoY': 'YOY',         # kolom Excel "Inflasi YoY" → sheet "YOY"
    'Andil MtM':   'AMTM',        # kolom Excel "Andil MtM"  → sheet "AMTM"
    'Andil YtD':   'AYTD',        # kolom Excel "Andil YtD"  → sheet "AYTD"
    'Andil YoY':   'AYOY',        # kolom Excel "Andil YoY"  → sheet "AYOY"
    'NK':   'NK',                 # kolom Excel "NK"  → sheet "NK"
}
#
# ──────────────────────────────────────────────────────────────────────────────

# Kolom tetap di setiap sheet Google Sheets (header row 3)
FIXED_COLS = ['Kd.Kota', 'Nama Kota', 'Kode', 'Nama Komoditas', 'Flag']

# Path credentials.json relatif terhadap root project
CREDS_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.dirname(__file__))),
    'credentials.json'
)

SCOPES = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]


# ─── Registry Dokumen (fetch dari GSheets dengan TTL cache) ───────────────────

import time as _time

_docs_cache: dict       = {}   # {section_key: list[dict]}
_docs_cache_ts: dict    = {}   # {section_key: float (unix timestamp)}


def get_docs(section: str, fallback: list) -> list:
    """
    Ambil daftar dokumen untuk section BRS dari Google Sheets registry.

    - section  : kunci dari DOCS_SHEET_MAP (misal 'IHK', 'NTP', ...)
    - fallback : list yang digunakan jika GSheets tidak bisa diakses.

    Data di-cache selama DOCS_CACHE_TTL detik.

    Format baris di Google Sheets (header baris 1, data mulai baris 2):
        Kolom A : Nama Dokumen
        Kolom B : Deskripsi
        Kolom C : Link
        Kolom D : Kategori  (misal: 'Master', 'BRS', 'Folder')

    Return — list of category groups:
        [{'kategori': 'Master', 'icon': 'bi-pencil-fill',
          'items': [{'label', 'desc', 'url', 'icon'}, ...]}, ...]
    """
    from app.constants import (
        DOCS_SPREADSHEET_ID, DOCS_SHEET_MAP, DOCS_CACHE_TTL,
    )

    if not DOCS_SPREADSHEET_ID:
        return fallback

    sheet_name = DOCS_SHEET_MAP.get(section)
    if not sheet_name:
        return fallback

    cache_key = section
    now       = _time.time()

    if cache_key in _docs_cache and (now - _docs_cache_ts.get(cache_key, 0)) < DOCS_CACHE_TTL:
        return _docs_cache[cache_key]

    # Icon per kategori (header group & item)
    _KAT_ICON = {
        'master' : 'bi-pencil-fill',
        'brs'    : 'bi-file-earmark-text-fill',
        'folder' : 'bi-folder2-open',
    }
    _ITEM_ICON = {
        'master' : 'bi-file-earmark-excel-fill',
        'brs'    : 'bi-file-earmark-text',
        'folder' : 'bi-folder',
    }

    try:
        spreadsheet = _get_spreadsheet(DOCS_SPREADSHEET_ID)
        worksheet = spreadsheet.worksheet(sheet_name)

        rows = worksheet.get_all_values()

        from collections import OrderedDict
        grouped = OrderedDict()

        for row in rows[1:]:
            if not any(row):
                continue
            label    = row[0].strip() if len(row) > 0 else ''
            desc     = row[1].strip() if len(row) > 1 else ''
            url      = row[2].strip() if len(row) > 2 else ''
            kategori = row[3].strip() if len(row) > 3 else 'Lainnya'

            if not label:
                continue

            kat_key   = kategori.lower()
            item_icon = _ITEM_ICON.get(kat_key, 'bi-box-arrow-up-right')

            if kategori not in grouped:
                grouped[kategori] = []
            grouped[kategori].append({
                'label': label,
                'desc' : desc,
                'url'  : url,
                'icon' : item_icon,
            })

        result = [
            {
                'kategori': kat,
                'icon'    : _KAT_ICON.get(kat.lower(), 'bi-collection'),
                'dokumen' : dok,
            }
            for kat, dok in grouped.items()
        ]

        _docs_cache[cache_key]    = result
        _docs_cache_ts[cache_key] = now
        return result

    except Exception as e:
        _log.error('get_docs(%s) error: %s', section, e, exc_info=True)
        return fallback


def clear_docs_cache(section: str = None):
    """
    Paksa refresh cache pada request berikutnya.
    Jika section=None, clear semua cache.
    """
    if section:
        _docs_cache.pop(section, None)
        _docs_cache_ts.pop(section, None)
    else:
        _docs_cache.clear()
        _docs_cache_ts.clear()


# ─── Klik Distribusi (fetch link portal dari GSheets) ─────────────────────────

_klik_cache: list = []
_klik_cache_ts: float = 0.0

_JENIS_ICON = {
    'internal app'  : 'bi-shield-lock',
    'dashboard'     : 'bi-speedometer2',
    'drive bps'     : 'bi-cloud-fill',
    'g-drive'       : 'bi-folder2-open',
    'g-sheets'      : 'bi-grid-3x3-gap-fill',
    'web app'       : 'bi-globe',
    'web entry'     : 'bi-keyboard',
    'website'       : 'bi-window',
}

_KATEGORI_COLOR = {
    'Utama'                        : 'primary',
    'Monitoring'                   : 'success',
    'Statistik Harga'              : 'danger',
    'Statistik Distribusi dan Jasa': 'purple',
    'Lainnya'                      : 'secondary',
}


def get_klik_links() -> list:
    """
    Fetch daftar link dari sheet 'Klik' di DOCS_SPREADSHEET_ID.
    Return list of dict per kategori:
      [{'kategori': 'SHK', 'color': 'danger', 'items': [{'nama', 'link', 'jenis', 'icon', 'keterangan'}, ...]}, ...]
    Di-cache selama DOCS_CACHE_TTL detik.
    Jika gagal, kembalikan list kosong.
    """
    from app.constants import KLIK_SPREADSHEET_ID, DOCS_CACHE_TTL

    global _klik_cache, _klik_cache_ts

    if not KLIK_SPREADSHEET_ID:
        return []

    now = _time.time()
    if _klik_cache and (now - _klik_cache_ts) < DOCS_CACHE_TTL:
        return _klik_cache

    try:
        spreadsheet = _get_spreadsheet(KLIK_SPREADSHEET_ID)
        worksheet = spreadsheet.worksheet('Klik')
        rows = worksheet.get_all_values()  # baris 1 = header
        if not rows:
            return []
            
        header = [h.strip().lower() for h in rows[0]]
        
        # Mapping index kolom secara dinamis
        def _get_idx(names: list[str], default: int) -> int:
            for name in names:
                if name in header:
                    return header.index(name)
            return default

        idx_kat  = _get_idx(['kategori', 'group'], 0)
        idx_nama = _get_idx(['nama', 'label', 'title'], 1)
        idx_link = _get_idx(['link', 'url'], 2)
        idx_jns  = _get_idx(['jenis', 'type'], 3)
        idx_kw   = _get_idx(['keyword', 'kata kunci'], 4)
        idx_ket  = _get_idx(['keterangan', 'desc', 'description'], 5)
        idx_thn  = _get_idx(['tahun', 'year'], 6)

        # Kumpulkan item per kategori (pertahankan urutan kemunculan)
        from collections import OrderedDict
        grouped = OrderedDict()

        for row in rows[1:]:               # skip header
            if not any(row):
                continue
            
            kategori    = row[idx_kat].strip() if len(row) > idx_kat else ''
            nama        = row[idx_nama].strip() if len(row) > idx_nama else ''
            link        = row[idx_link].strip() if len(row) > idx_link else '#'
            jenis       = row[idx_jns].strip() if len(row) > idx_jns else ''
            keyword     = row[idx_kw].strip() if len(row) > idx_kw else ''
            keterangan  = row[idx_ket].strip() if len(row) > idx_ket else ''
            tahun_raw   = row[idx_thn].strip() if len(row) > idx_thn else ''

            if not nama:
                continue

            # Konversi tahun ke int jika memungkinkan
            try:
                tahun = int(float(tahun_raw)) if tahun_raw else None
            except (ValueError, TypeError):
                tahun = None

            icon = _JENIS_ICON.get(jenis.lower(), 'bi-box-arrow-up-right')

            if kategori not in grouped:
                grouped[kategori] = []
            grouped[kategori].append({
                'nama'      : nama,
                'link'      : link,
                'jenis'     : jenis,
                'icon'      : icon,
                'keyword'   : keyword,
                'keterangan': keterangan,
                'tahun'     : tahun,
            })

        result = [
            {
                'kategori': kat,
                'color'   : _KATEGORI_COLOR.get(kat, 'secondary'),
                'links'   : links,
            }
            for kat, links in grouped.items()
        ]

        _klik_cache    = result
        _klik_cache_ts = now
        return result

    except Exception:
        return _klik_cache or []   # kembalikan cache lama jika ada, atau kosong


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

def _upsert_ihk_sheet(
    worksheet: gspread.Worksheet,
    dataframe,
    col_name: str,
    excel_col_name: str,
) -> tuple[int, int]:
    """
    Insert/update data dari DataFrame ke worksheet IHK `worksheet`.
    
    Args:
        worksheet: Gspread Worksheet instance
        dataframe: Pandas DataFrame containing IHK values
        col_name: kode YYMM (misal '2608') untuk header kolom
        excel_col_name: nama kolom asal di DataFrame excel (misal 'IHK', 'Inflasi MtM', dst.)
        
    Returns:
        tuple (inserted_rows_count, updated_rows_count)
    """
    all_values = worksheet.get_all_values()
    header_row = all_values[2] if len(all_values) > 2 else []
    col_indices = {name: _col_index(header_row, name) for name in FIXED_COLS}

    # Cari/buat kolom YYMM dan ambil data terbaru
    val_col_1based, is_new = find_or_create_col(worksheet, header_row, col_name)
    
    if is_new:
        # Tentukan nomor urut dari baris ke-4
        row4 = all_values[3] if len(all_values) > 3 else []
        num_filled = sum(1 for x in row4 if str(x).strip())
        next_num = num_filled + 1
        
        worksheet.batch_update([
            {'range': rowcol_to_a1(3, val_col_1based), 'values': [[col_name]]},
            {'range': rowcol_to_a1(4, val_col_1based), 'values': [[next_num]]},
        ])
        # Update LOCAL representation agar tidak perlu get_all_values() lagi
        # Kita hanya perlu memastikan baris data yang diambil benar.
        # Karena kita hanya tambah KOLOM di header, baris data tidak bergeser.
        # Namun untuk safety, kita tambahkan dummy value ke header agar index valid.
        while len(header_row) < val_col_1based:
            header_row.append('')
        header_row[val_col_1based-1] = col_name

    # all_values = worksheet.get_all_values() # DIHAPUS: Mengurangi 1 Read Request per sheet

    data_rows = all_values[4:] 
    lookup = _build_row_lookup(data_rows, col_indices)

    batch_updates = []
    new_rows = []
    inserted_count = 0
    updated_count = 0

    max_col_0 = max((v for v in col_indices.values() if v is not None), default=0)
    row_width = max(max_col_0, val_col_1based - 1) + 1

    for _, row in dataframe.iterrows():
        kd_kota = str(row.get('Kode Kota', '')).strip()
        kode_komoditas = str(row.get('Kode Komoditas', '')).strip()
        cell_value = _to_number(row.get(excel_col_name, ''))

        key = (kd_kota, kode_komoditas)

        if key in lookup:
            row_idx = lookup[key]
            batch_updates.append({
                'range': rowcol_to_a1(row_idx, val_col_1based),
                'values': [[cell_value]],
            })
            updated_count += 1
        else:
            row_data = [''] * row_width
            
            def _set(col_key: str, val: str):
                idx = col_indices.get(col_key)
                if idx is not None and idx < row_width:
                    row_data[idx] = val

            _set('Kd.Kota', kd_kota)
            _set('Nama Kota', str(row.get('Nama Kota', '')).strip())
            _set('Kode', kode_komoditas)
            _set('Nama Komoditas', str(row.get('Nama Komoditas', '')).strip())
            _set('Flag', str(row.get('Flag', '')).strip())
            
            if val_col_1based - 1 < row_width:
                row_data[val_col_1based - 1] = cell_value

            new_rows.append(row_data)
            next_row = 5 + len(data_rows) + len(new_rows) - 1
            lookup[key] = next_row
            inserted_count += 1

    if batch_updates:
        worksheet.batch_update(batch_updates)

    if new_rows:
        next_empty_row = max(len(all_values) + 1, 5)
        last_needed_row = next_empty_row + len(new_rows) - 1

        if last_needed_row > worksheet.row_count:
            worksheet.add_rows(last_needed_row - worksheet.row_count + 200)

        # Optimasi: Gunakan satu range update alih-alih batch_update baris demi baris
        range_name = f"{rowcol_to_a1(next_empty_row, 1)}:{rowcol_to_a1(last_needed_row, len(new_rows[0]))}"
        worksheet.update(new_rows, range_name)

    return inserted_count, updated_count

# ─── IHK Orchestrator ─────────────────────────────────────────────────────────────

def process_ihk_upload(dataframe, col_name: str, task_id: str = None) -> dict:
    """
    Upload IHK DataFrame ke 7 sheet Google Sheets sekaligus.
    Return dict {'inserted': int, 'updated': int, 'details': list}.
    """
    set_progress(task_id, 10, "Menyambungkan ke Google Sheets API...")
    spreadsheet = _get_spreadsheet(IHK_SPREADSHEET_ID)

    total_inserted = 0
    total_updated  = 0
    details = []

    num_sheets = len(SHEET_VALUE_MAP)
    set_progress(task_id, 20, f"Mempersiapkan {num_sheets} sheet...")

    for idx, (excel_col, sheet_name) in enumerate(SHEET_VALUE_MAP.items(), 1):
        msg = f"Memproses sheet '{sheet_name}' ({idx}/{num_sheets})..."
        pct = 20 + int(80 * (idx - 1) / num_sheets)
        set_progress(task_id, pct, msg)
        
        worksheet = spreadsheet.worksheet(sheet_name)
        # Tambahkan retry wrapper agar tidak gagal total jika kena limit
        ins, upd = _retry_on_429(_upsert_ihk_sheet)(worksheet, dataframe, col_name, excel_col, _task_id=task_id)
        total_inserted += ins
        total_updated  += upd
        details.append({'sheet': sheet_name, 'inserted': ins, 'updated': upd})
        
        # Beri jeda antar sheet untuk meminimalkan resiko rate limit
        time.sleep(1)

    set_progress(task_id, 100, "Selesai!")
    return {
        'inserted': total_inserted,
        'updated':  total_updated,
        'details':  details,
    }

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
        try:
            yymm_val = int(yymm)
        except:
            yymm_val = yymm
            
        if yymm in lookup:
            row_idx = lookup[yymm]
            batch_updates.append({'range': rowcol_to_a1(row_idx, col_map['Migas']), 'values': [[vals.get('MIGAS', 0)]]})
            batch_updates.append({'range': rowcol_to_a1(row_idx, col_map['Non Migas']), 'values': [[vals.get('NON MIGAS', 0)]]})
            batch_updates.append({'range': rowcol_to_a1(row_idx, col_map['Total']), 'values': [[vals.get('Total', 0)]]})
        else:
            row_data = [''] * len(header_row)
            row_data[col_map['Periode']-1] = yymm_val
            row_data[col_map['Migas']-1] = vals.get('MIGAS', 0)
            row_data[col_map['Non Migas']-1] = vals.get('NON MIGAS', 0)
            row_data[col_map['Total']-1] = vals.get('Total', 0)
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

def _upsert_ekspor_sheet_bulk(worksheet: gspread.Worksheet, months_data: dict, key_col_name: str | list[str], sheet_name: str = None, upload_type: str = 'ekspor') -> tuple[int, int]:
    """
    Menambahkan atau mem-patch aggregated FOB sum ke Google Sheet Ekspor secara BULK.
    months_data: { yymm_label: dataframe_group }
    key_col_name: string (single key) atau list of strings (composite keys)
    """
    all_values = worksheet.get_all_values() # HANYA 1x READ CALL per sheet
    header_row = all_values[0] if len(all_values) > 0 else []
    
    key_col_names = [key_col_name] if isinstance(key_col_name, str) else key_col_name
    
    prov_col_name = 'KodeProvTujuan' if upload_type == 'impor' else 'KodeProvAsal'

    if not header_row:
        # Jika sheet kosong, inisialisasi header dasar
        sample_yymm = list(months_data.keys())[0] if months_data else 'YYMM'
        if sheet_name == 'Pelabuhan':
            header_row = key_col_names + [prov_col_name, sample_yymm]
        else:
            header_row = key_col_names + [sample_yymm]
        worksheet.append_row(header_row)
        all_values = [header_row]
        
    # Ambil index untuk semua key columns
    key_indices = []
    for kn in key_col_names:
        try:
            key_indices.append(header_row.index(kn))
        except ValueError:
            # Jika kolom tidak ada di header, kita tambahkan manual ke header (lokal)
            # dan nanti akan ter-update via batch_header_updates
            key_indices.append(len(header_row))
            header_row.append(kn)

    # Khusus Pelabuhan: Cek/Sisipkan KodeProvAsal/KodeProvTujuan tepat di kanan key pertama
    prov_col_idx = None
    if sheet_name == 'Pelabuhan':
        primary_key_idx = key_indices[0]
        if prov_col_name in header_row:
            prov_col_idx = header_row.index(prov_col_name)
        else:
            prov_col_idx = primary_key_idx + 1
            # Sisipkan kolom via API (one-time operation)
            _log.info(f"Menyisipkan kolom {prov_col_name} ke sheet {sheet_name}...")
            worksheet.insert_cols([[prov_col_name]], prov_col_idx + 1)
            # Refetch data karena semua kolom bergeser
            all_values = worksheet.get_all_values()
            header_row = all_values[0]
            prov_col_idx = header_row.index(prov_col_name)
            # Reset key_indices karena kolom bergeser
            key_indices = [header_row.index(kn) for kn in key_col_names]

    # Khusus NegaraBarang: Tidak ada lagi metadata BRSDESC, OIL, KEL_IMPOR
    # Hanya pastikan HS key teridentifikasi jika masih dibutuhkan (biasanya tidak jika hanya 2 kolom keys)
    
    data_rows = all_values[1:]
    
    def _get_lookup_key(row_data, idx_list):
        return tuple(str(row_data[i]).strip() if i < len(row_data) else '' for i in idx_list)

    lookup = {}
    for i, row in enumerate(data_rows, start=2):
        k = _get_lookup_key(row, key_indices)
        if any(k): # Minimal ada satu bagian key yang tidak kosong
            lookup[k] = i
            
    batch_updates = []
    batch_header_updates = []
    new_rows_map = {} # key_val -> [col1, col2, ...]
    
    inserted_count = 0
    updated_count = 0
    
    # Proses setiap bulan
    for yymm_label, df_group in months_data.items():
        yymm_idx = _ensure_ekspor_yymm_column(worksheet, header_row, yymm_label, batch_header_updates)
        # Re-calculate row_width
        row_width = max(len(header_row), (max(key_indices) + 1) if key_indices else 0)
        if prov_col_idx is not None:
            row_width = max(row_width, prov_col_idx + 1)
        
        
        for _, row in df_group.iterrows():
            # Build current key as tuple
            curr_key_parts = []
            for kn in key_col_names:
                curr_key_parts.append(str(row.get(kn, '')).strip())
            curr_key = tuple(curr_key_parts)
            
            val_fob = float(row['FOB'])
            
            if curr_key in lookup:
                row_idx = lookup[curr_key]
                batch_updates.append({
                    'range': rowcol_to_a1(row_idx, yymm_idx),
                    'values': [[val_fob]]
                })
                # Opsional: Pastikan KodeProvAsal/Tujuan terisi jika kosong (khusus Pelabuhan)
                if prov_col_idx is not None:
                    curr_row_idx = row_idx - 1
                    curr_row = all_values[curr_row_idx] if curr_row_idx < len(all_values) else []
                    if prov_col_idx >= len(curr_row) or not str(curr_row[prov_col_idx]).strip():
                        batch_updates.append({
                            'range': rowcol_to_a1(row_idx, prov_col_idx + 1),
                            # Ambil 2 digit dari key pertama (PODAL5 / K_PELB)
                            'values': [[curr_key[0][:2]]]
                        })
                
                updated_count += 1
            else:
                # Handle baris baru
                if curr_key not in new_rows_map:
                    row_data = [''] * row_width
                    # Set all key columns
                    for i, idx in enumerate(key_indices):
                        if idx < row_width:
                            row_data[idx] = curr_key[i]
                    
                    if prov_col_idx is not None and prov_col_idx < row_width:
                        row_data[prov_col_idx] = curr_key[0][:2]
                        
                    new_rows_map[curr_key] = row_data
                    inserted_count += 1
                
                # Update kolom bulan di baris baru
                if yymm_idx > len(new_rows_map[curr_key]):
                    new_rows_map[curr_key].extend([''] * (yymm_idx - len(new_rows_map[curr_key])))
                new_rows_map[curr_key][yymm_idx - 1] = val_fob

    # Eksekusi Updates
    if batch_header_updates:
        worksheet.batch_update(batch_header_updates)
        
    if batch_updates:
        # Pecah batch updates jika terlalu besar untuk menghindari error lain
        # (meskipun biasanya limitnya cukup besar untuk cell updates)
        worksheet.batch_update(batch_updates)
        
    if new_rows_map:
        new_rows = list(new_rows_map.values())
        # Pastikan semua baris memiliki panjang yang sama (padding) agar tidak error saat batch update
        max_cols_in_new = max(len(r) for r in new_rows)
        for i in range(len(new_rows)):
            diff = max_cols_in_new - len(new_rows[i])
            if diff > 0:
                new_rows[i].extend([''] * diff)

        next_empty_row = max(len(all_values) + 1, 2)
        last_needed_row = next_empty_row + len(new_rows) - 1
        
        if last_needed_row > worksheet.row_count:
            worksheet.add_rows(last_needed_row - worksheet.row_count + 50)
            
        range_name = f"{rowcol_to_a1(next_empty_row, 1)}:{rowcol_to_a1(last_needed_row, max_cols_in_new)}"
        worksheet.update(new_rows, range_name)
        
    # ── SORTING AKHIR (Rows & Columns) ──────────────────────────────────────
    try:
        # 1. Sort Rows by Keys (Composite)
        final_row_count = worksheet.row_count
        if final_row_count > 1:
            sort_specs = [{"dimensionIndex": idx, "sortOrder": "ASCENDING"} for idx in key_indices]
            body = {
                "requests": [{
                    "sortRange": {
                        "range": {
                            "sheetId": worksheet.id,
                            "startRowIndex": 1,
                            "endRowIndex": final_row_count,
                            "startColumnIndex": 0,
                            "endColumnIndex": len(header_row)
                        },
                        "sortSpecs": sort_specs
                    }
                }]
            }
            worksheet.spreadsheet.batch_update(body)

        # 2. Horizontal Sorting (Columns)
        curr_header = worksheet.get_all_values()[0]
        periods = []
        for i, val in enumerate(curr_header):
            v = str(val).strip()
            if v.isdigit() and len(v) == 4:
                periods.append((v, i))
        
        if periods:
            # Urutkan periods berdasarkan label YYMM secara ascending
            sorted_periods = sorted(periods, key=lambda x: x[0])
            
            # Cek apakah sudah terurut
            current_indices = [p[1] for p in periods]
            # Kita tidak bisa sekadar bandingkan indices karena positions bisa geser
            # Kita cek apakah labels di curr_header sudah terurut di posisinya
            
            # Simulasi untuk batch move
            move_requests = []
            header_sim = list(curr_header)
            # Tentukan start offset (posisi kolom pertama yang berisi periode)
            # Namun agar lebih stabil, kita ambil index terkecil dari kolom periode yang ada
            start_offset = min(p[1] for p in periods)
            
            for target_pos, (label, _) in enumerate(sorted_periods):
                final_target_idx = start_offset + target_pos
                try:
                    actual_idx = header_sim.index(label)
                except ValueError:
                    continue
                
                if actual_idx != final_target_idx:
                    move_requests.append({
                        "moveDimension": {
                            "source": {
                                "sheetId": worksheet.id,
                                "dimension": "COLUMNS",
                                "startIndex": actual_idx,
                                "endIndex": actual_idx + 1
                            },
                            "destinationIndex": final_target_idx if final_target_idx < actual_idx else final_target_idx + 1
                        }
                    })
                    # Update simulasi
                    col = header_sim.pop(actual_idx)
                    header_sim.insert(final_target_idx, col)
            
            if move_requests:
                _log.info(f"Mengurutkan {len(move_requests)} kolom secara horizontal di {sheet_name}...")
                worksheet.spreadsheet.batch_update({"requests": move_requests})

    except Exception as e:
        _log.error(f"Gagal sorting akhir {sheet_name}: {e}")
        
    return inserted_count, updated_count

def process_ekspor_upload(dataframe, task_id: str = None) -> dict:
    """Menerima dataframe ekspor, mengelompokkan jumlah FOB secara bulk, dan menguploadnya."""
    import pandas as pd
    
    set_progress(task_id, 10, "Menyambungkan ke Google Sheets API...")
    spreadsheet = _get_spreadsheet(EKSPOR_IMPOR_SPREADSHEET_ID)
    
    dataframe['FOB'] = pd.to_numeric(dataframe['FOB'], errors='coerce').fillna(0)
    
    if 'YEAR' not in dataframe.columns or 'MTH' not in dataframe.columns:
        raise ValueError("Kolom 'YEAR' atau 'MTH' tidak ditemukan di data.")
        
    for req_col in ['KODE_HS', 'PODAL5', 'NEWCTRYCOD']:
        if req_col not in dataframe.columns:
            raise ValueError(f"Kolom '{req_col}' tidak ditemukan di data.")
    
    total_inserted = 0
    total_updated = 0
    details = []
    
    # 1.5 Fetch HS Master map
    set_progress(task_id, 20, "Mengambil data MASTER referensi...")
    hs_master_map = _get_master_hs_data(spreadsheet)
    
    # 1. Pre-group data by month
    set_progress(task_id, 30, "Mempersiapkan pengelompokan baris per bulan...")
    months_payload = {}
    for (year, month), group in dataframe.groupby(['YEAR', 'MTH']):
        if not str(year).strip() or not str(month).strip():
            continue
        try:
            year_str = str(int(float(year)))[-2:]
            month_str = str(int(float(month))).zfill(2)
            yymm_label = year_str + month_str
            months_payload[yymm_label] = group
        except (ValueError, TypeError):
            continue

    if not months_payload:
        return {'inserted': 0, 'updated': 0, 'details': [{'error': 'No valid month/year data found'}]}

    # 2. Process per sheet (Total only 2 sheets for Ekspor)
    targets = [
        ('Pelabuhan', 'PODAL5'),
        ('NegaraBarang', ['NEWCTRYCOD', 'KODE_HS'])
    ]
    
    for idx, (target_sheet_name, group_key_col) in enumerate(targets, 1):
        def _execute_bulk():
            # Agregasi data untuk setiap bulan khusus untuk sheet ini
            set_progress(task_id, 40 + (idx-1)*20, f"Agregasi data untuk sheet '{target_sheet_name}'...")
            prepared_months = {}
            for yymm, group in months_payload.items():
                summed = group.groupby(group_key_col)['FOB'].sum().reset_index()
                prepared_months[yymm] = summed
            
            worksheet = spreadsheet.worksheet(target_sheet_name)
            return _upsert_ekspor_sheet_bulk(worksheet, prepared_months, group_key_col, target_sheet_name, upload_type='ekspor')
            
        ins, upd = _retry_on_429(_execute_bulk)()
        total_inserted += ins
        total_updated += upd
        details.append({'sheet': target_sheet_name, 'inserted': ins, 'updated': upd})
        
        # Jeda antar sheet
        time.sleep(2)

    # 3. Process RECAP sheet
    def _execute_recap():
        set_progress(task_id, 90, "Agregasi data untuk sheet 'RECAP'...")
        # Map OIL and aggregate
        df_recap = dataframe.copy()
        df_recap['OIL'] = df_recap['KODE_HS'].map(lambda x: hs_master_map.get(str(x).strip(), 'NON MIGAS'))
        # Normalize OIL names to MIGAS or NON MIGAS
        df_recap['OIL'] = df_recap['OIL'].apply(lambda x: 'MIGAS' if 'MIGAS' in str(x).upper() and 'NON' not in str(x).upper() else 'NON MIGAS')
        
        recap_payload = {}
        for (year, month), group in df_recap.groupby(['YEAR', 'MTH']):
            try:
                yymm = str(int(float(year)))[-2:] + str(int(float(month))).zfill(2)
                oil_sums = group.groupby('OIL')['FOB'].sum().to_dict()
                total = group['FOB'].sum()
                recap_payload[yymm] = {
                    'MIGAS': oil_sums.get('MIGAS', 0),
                    'NON MIGAS': oil_sums.get('NON MIGAS', 0),
                    'Total': total
                }
            except:
                continue
                
        if recap_payload:
            worksheet = spreadsheet.worksheet('RECAP')
            ins, upd = _upsert_recap_sheet(worksheet, recap_payload)
            return ins, upd
        return 0, 0

    try:
        ins_r, upd_r = _retry_on_429(_execute_recap)()
        details.append({'sheet': 'RECAP', 'inserted': ins_r, 'updated': upd_r})
        total_inserted += ins_r
        total_updated += upd_r
    except Exception as e:
        _log.error(f"Gagal update RECAP: {e}")
        details.append({'sheet': 'RECAP', 'error': str(e)})

    set_progress(task_id, 100, "Selesai!")
    return {
        'inserted': total_inserted,
        'updated': total_updated,
        'details': details
    }


# ─── Impor Upload (Similar to Ekspor) ─────────────────────────────────────────

def process_impor_upload(dataframe, task_id: str = None) -> dict:
    """Menerima dataframe impor, mengelompokkan jumlah N1225/NILAI secara bulk, dan menguploadnya."""
    import pandas as pd
    
    set_progress(task_id, 10, "Menyambungkan ke Google Sheets API...")
    spreadsheet = _get_spreadsheet(IMPOR_SPREADSHEET_ID)
    
    dataframe['N1225'] = pd.to_numeric(dataframe['N1225'], errors='coerce').fillna(0)
    
    if 'YEAR' not in dataframe.columns or 'MTH' not in dataframe.columns:
        raise ValueError("Kolom 'YEAR' atau 'MTH' tidak ditemukan di data.")
        
    for req_col in ['HS', 'K_NEGARA']:
        if req_col not in dataframe.columns:
            raise ValueError(f"Kolom '{req_col}' tidak ditemukan di data.")
    
    total_inserted = 0
    total_updated = 0
    details = []
    
    # 1.5 Fetch HS Master map
    set_progress(task_id, 20, "Mengambil data MASTER referensi...")
    hs_master_map = _get_master_hs_data(spreadsheet)
    
    # 1. Pre-group data by month
    set_progress(task_id, 30, "Mempersiapkan pengelompokan baris per bulan...")
    months_payload = {}
    for (year, month), group in dataframe.groupby(['YEAR', 'MTH']):
        if not str(year).strip() or not str(month).strip():
            continue
        try:
            year_str = str(int(float(year)))[-2:]
            month_str = str(int(float(month))).zfill(2)
            yymm_label = year_str + month_str
            months_payload[yymm_label] = group
        except (ValueError, TypeError):
            continue

    if not months_payload:
        return {'inserted': 0, 'updated': 0, 'details': [{'error': 'No valid month/year data found'}]}

    # 2. Process per sheet
    targets = [
        ('NegaraBarang', ['K_NEGARA', 'HS'])
    ]
    
    for idx, (target_sheet_name, group_key_col) in enumerate(targets, 1):
        def _execute_bulk():
            set_progress(task_id, 40 + (idx-1)*20, f"Agregasi data untuk sheet '{target_sheet_name}'...")
            prepared_months = {}
            for yymm, group in months_payload.items():
                summed = group.groupby(group_key_col)['N1225'].sum().reset_index()
                # _upsert_ekspor_sheet_bulk expects 'FOB' col
                summed = summed.rename(columns={'N1225': 'FOB'})
                prepared_months[yymm] = summed
            
            worksheet = spreadsheet.worksheet(target_sheet_name)
            return _upsert_ekspor_sheet_bulk(worksheet, prepared_months, group_key_col, target_sheet_name, upload_type='impor')
            
        ins, upd = _retry_on_429(_execute_bulk)()
        total_inserted += ins
        total_updated += upd
        details.append({'sheet': target_sheet_name, 'inserted': ins, 'updated': upd})
        
        time.sleep(2)

    # 3. Process RECAP sheet
    def _execute_recap():
        set_progress(task_id, 90, "Agregasi data untuk sheet 'RECAP'...")
        # Map OIL and aggregate
        df_recap = dataframe.copy()
        df_recap['OIL'] = df_recap['HS'].map(lambda x: hs_master_map.get(str(x).strip(), 'NON MIGAS'))
        # Normalize OIL names to MIGAS or NON MIGAS
        df_recap['OIL'] = df_recap['OIL'].apply(lambda x: 'MIGAS' if 'MIGAS' in str(x).upper() and 'NON' not in str(x).upper() else 'NON MIGAS')
        
        recap_payload = {}
        for (year, month), group in df_recap.groupby(['YEAR', 'MTH']):
            try:
                yymm = str(int(float(year)))[-2:] + str(int(float(month))).zfill(2)
                oil_sums = group.groupby('OIL')['N1225'].sum().to_dict()
                total = group['N1225'].sum()
                recap_payload[yymm] = {
                    'MIGAS': oil_sums.get('MIGAS', 0),
                    'NON MIGAS': oil_sums.get('NON MIGAS', 0),
                    'Total': total
                }
            except:
                continue
                
        if recap_payload:
            worksheet = spreadsheet.worksheet('RECAP')
            ins, upd = _upsert_recap_sheet(worksheet, recap_payload)
            return ins, upd
        return 0, 0

    try:
        ins_r, upd_r = _retry_on_429(_execute_recap)()
        details.append({'sheet': 'RECAP', 'inserted': ins_r, 'updated': upd_r})
        total_inserted += ins_r
        total_updated += upd_r
    except Exception as e:
        _log.error(f"Gagal update RECAP: {e}")
        details.append({'sheet': 'RECAP', 'error': str(e)})

    set_progress(task_id, 100, "Selesai!")
    return {
        'inserted': total_inserted,
        'updated': total_updated,
        'details': details
    }
