"""
Google Sheets Service — IHK Upload
Mengelola koneksi dan operasi tulis ke Google Sheets via service account.
"""

import os
import logging

_log = logging.getLogger(__name__)
import gspread
from gspread.utils import rowcol_to_a1
from google.oauth2.service_account import Credentials

# ─── Konstanta ────────────────────────────────────────────────────────────────

SPREADSHEET_ID = '1iB0DMQw7kjVzl5PgrVsb1wtQIlRF3QL-pZHYmQ0JySQ'

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
        gc = get_client()
        ss = gc.open_by_key(DOCS_SPREADSHEET_ID)
        ws = ss.worksheet(sheet_name)

        rows = ws.get_all_values()

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
    'google sheets' : 'bi-grid-3x3-gap-fill',
    'web app'       : 'bi-globe',
    'web entry'     : 'bi-keyboard',
    'upload form'   : 'bi-cloud-upload',
    'dashboard'     : 'bi-speedometer2',
    'website'       : 'bi-window',
    'internal app'  : 'bi-shield-lock',
    'drive/cloud'   : 'bi-cloud-fill',
    'google drive'  : 'bi-folder2-open',
    'shortlink'     : 'bi-link-45deg',
    'upload'        : 'bi-arrow-up-circle',
    'upload/dashboard': 'bi-arrow-up-circle',
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
        gc   = get_client()
        ss   = gc.open_by_key(KLIK_SPREADSHEET_ID)
        ws   = ss.worksheet('Klik')
        rows = ws.get_all_values()  # baris 1 = header

        # Kumpulkan item per kategori (pertahankan urutan kemunculan)
        from collections import OrderedDict
        grouped = OrderedDict()

        for row in rows[1:]:               # skip header
            if not any(row):
                continue
            kategori    = row[0].strip() if len(row) > 0 else ''
            nama        = row[1].strip() if len(row) > 1 else ''
            link        = row[2].strip() if len(row) > 2 else '#'
            jenis       = row[3].strip() if len(row) > 3 else ''
            keyword     = row[4].strip() if len(row) > 4 else ''
            keterangan  = row[5].strip() if len(row) > 5 else ''

            if not nama:
                continue

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

def get_client() -> gspread.Client:
    """
    Buat koneksi gspread menggunakan service account.

    - Di Vercel/production: baca dari env var GOOGLE_CREDENTIALS_JSON (JSON string)
    - Di local dev: fallback ke file credentials.json
    """
    creds_json = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    if creds_json:
        import json
        info = json.loads(creds_json)
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file(CREDS_PATH, scopes=SCOPES)
    return gspread.authorize(creds)


# ─── Helper ───────────────────────────────────────────────────────────────────

def _col_index(header_row: list, name: str) -> int | None:
    """Kembalikan 0-based index kolom dari header_row, atau None jika tidak ada."""
    try:
        return header_row.index(name)
    except ValueError:
        return None


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


def find_or_create_col(ws: gspread.Worksheet, all_vals: list, col_name: str) -> int:
    """
    Cari kolom `col_name` di baris ke-3 (index 2).
    Jika tidak ada, buat kolom baru di ujung kanan:
      - Baris 3: tulis col_name
      - Baris 4: tulis nomor urut
    Kembalikan index kolom (1-based).
    """
    header_row = all_vals[2] if len(all_vals) > 2 else []

    if col_name in header_row:
        return header_row.index(col_name) + 1  # 1-based

    # Kolom baru
    new_col_1based = len(header_row) + 1

    # Tentukan nomor urut dari baris ke-4
    row4 = all_vals[3] if len(all_vals) > 3 else []
    num_filled = len([x for x in row4 if str(x).strip()])
    next_num = num_filled + 1

    ws.batch_update([
        {'range': rowcol_to_a1(3, new_col_1based), 'values': [[col_name]]},
        {'range': rowcol_to_a1(4, new_col_1based), 'values': [[next_num]]},
    ])

    return new_col_1based


# ─── Per-sheet Logic ──────────────────────────────────────────────────────────

def process_sheet(
    ws: gspread.Worksheet,
    df,
    col_name: str,
    excel_col: str,
) -> tuple[int, int]:
    """
    Insert/update data dari DataFrame ke worksheet `ws`.

    - col_name : kode YYMM (misal '2608')
    - excel_col: nama kolom di Excel (misal 'IHK', 'Inflasi MtM', dst.)

    Return (inserted, updated).
    """
    # Ambil seluruh data sheet sekali saja
    all_vals = ws.get_all_values()

    header_row = all_vals[2] if len(all_vals) > 2 else []  # baris 3 (0-based index 2)

    # Indeks kolom tetap (0-based)
    ci = {name: _col_index(header_row, name) for name in FIXED_COLS}

    # Cari/buat kolom YYMM
    val_col_1based = find_or_create_col(ws, all_vals, col_name)

    # Refresh all_vals jika kolom baru dibuat
    all_vals = ws.get_all_values()

    # Bangun lookup: (kd_kota, kode) → baris ke-N (1-based, ≥5)
    data_rows = all_vals[4:]  # baris 5 dst (0-based index 4)
    lookup: dict[tuple, int] = {}
    for i, row in enumerate(data_rows, start=5):
        kd  = row[ci['Kd.Kota']] if ci['Kd.Kota'] is not None and ci['Kd.Kota'] < len(row) else ''
        kode = row[ci['Kode']]   if ci['Kode']    is not None and ci['Kode']    < len(row) else ''
        if kd or kode:
            lookup[(str(kd).strip(), str(kode).strip())] = i

    # Proses baris Excel
    batch_updates: list[dict] = []
    new_rows: list[list] = []
    inserted, updated = 0, 0

    # Tentukan lebar baris baru
    max_col_0 = max(
        (v for v in ci.values() if v is not None),
        default=0
    )
    max_col_0 = max(max_col_0, val_col_1based - 1)
    row_width = max_col_0 + 1

    for _, er in df.iterrows():
        kd_kota  = str(er.get('Kode Kota', '')).strip()
        kode     = str(er.get('Kode Komoditas', '')).strip()
        value    = _to_number(er.get(excel_col, ''))  # simpan sebagai numerik
        nama_kota = str(er.get('Nama Kota', '')).strip()
        nama_kom  = str(er.get('Nama Komoditas', '')).strip()
        flag      = str(er.get('Flag', '')).strip()

        key = (kd_kota, kode)

        if key in lookup:
            # Update nilai di kolom YYMM
            row_idx = lookup[key]
            batch_updates.append({
                'range': rowcol_to_a1(row_idx, val_col_1based),
                'values': [[value]],
            })
            updated += 1
        else:
            # Baris baru
            row_data = [''] * row_width
            def _set(col_name_key, val):
                idx = ci.get(col_name_key)
                if idx is not None and idx < row_width:
                    row_data[idx] = val

            _set('Kd.Kota', kd_kota)
            _set('Nama Kota', nama_kota)
            _set('Kode', kode)
            _set('Nama Komoditas', nama_kom)
            _set('Flag', flag)
            if val_col_1based - 1 < row_width:
                row_data[val_col_1based - 1] = value

            new_rows.append(row_data)
            # Daftarkan ke lookup agar tidak duplikat dalam satu upload
            next_row = 5 + len(data_rows) + len(new_rows) - 1
            lookup[key] = next_row
            inserted += 1

    # Eksekusi batch update (update baris lama)
    if batch_updates:
        ws.batch_update(batch_updates)

    # Append baris baru via batch_update
    if new_rows:
        next_empty_row = max(len(all_vals) + 1, 5)
        last_needed_row = next_empty_row + len(new_rows) - 1

        # Pastikan sheet punya cukup baris (auto-expand jika perlu)
        if last_needed_row > ws.row_count:
            rows_to_add = last_needed_row - ws.row_count + 200  # tambah buffer 200
            ws.add_rows(rows_to_add)

        append_batch = []
        for i, row_data in enumerate(new_rows):
            r = next_empty_row + i
            end_col = len(row_data)
            rng = f"{rowcol_to_a1(r, 1)}:{rowcol_to_a1(r, end_col)}"
            append_batch.append({'range': rng, 'values': [row_data]})
        ws.batch_update(append_batch)

    return inserted, updated


# ─── Orchestrator ─────────────────────────────────────────────────────────────

def process_upload(df, col_name: str) -> dict:
    """
    Upload DataFrame ke 7 sheet Google Sheets sekaligus.
    Return dict {'inserted': int, 'updated': int, 'details': list}.
    """
    gc = get_client()
    ss = gc.open_by_key(SPREADSHEET_ID)

    total_inserted = 0
    total_updated  = 0
    details = []

    for excel_col, sheet_name in SHEET_VALUE_MAP.items():
        ws = ss.worksheet(sheet_name)
        ins, upd = process_sheet(ws, df, col_name, excel_col)
        total_inserted += ins
        total_updated  += upd
        details.append({'sheet': sheet_name, 'inserted': ins, 'updated': upd})

    return {
        'inserted': total_inserted,
        'updated':  total_updated,
        'details':  details,
    }
