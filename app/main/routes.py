"""
Routes untuk blueprint main (halaman utama dan BRS).
"""
from flask import render_template, request, redirect, url_for, flash, send_file
from flask_login import login_required
from app.main import main
from app.services.sheets import get_docs, get_klik_links, clear_docs_cache
import app.services.sheets as _sheets
import io



def _safe_int(val):
    """Konversi nilai ke int dengan aman — menangani string, float, dan NaN."""
    try:
        return int(float(str(val).strip()))
    except (ValueError, TypeError):
        return None



@main.route('/')
@main.route('/index')
def index():
    klik_groups = get_klik_links()
    return render_template('index.html', title='Klik Distribusi', klik_groups=klik_groups)


@main.route('/klik-refresh')
def klik_refresh():
    """Hapus cache Klik Distribusi lalu redirect ke halaman utama."""
    from app.services.sheets import _klik_cache, _klik_cache_ts
    import app.services.sheets as _sheets
    _sheets._klik_cache    = []
    _sheets._klik_cache_ts = 0.0
    flash('Klik Distribusi telah di-refresh — Daftar link berhasil diperbarui.', 'success')
    return redirect(url_for('main.index'))


@main.route('/docs-refresh')
@login_required
def docs_refresh():
    """Hapus cache semua BRS docs lalu redirect kembali ke halaman asal."""
    import app.services.sheets as _sheets
    _sheets._docs_cache.clear()
    _sheets._docs_cache_ts.clear()
    flash('Dokumen-dokumen BRS telah di-refresh — Daftar dokumen berhasil diperbarui.', 'success')
    referrer = request.referrer
    if referrer:
        return redirect(referrer)
    return redirect(url_for('main.dia_brs'))


@main.route('/dashboard')
@login_required
def dashboard():
    return redirect(url_for('main.dia_brs'))


@main.route('/dia-brs')
@login_required
def dia_brs():
    return render_template('dashboard.html', title='DIA BRS — Dashboard')


from app.services.sheets import get_docs


@main.route('/brs/ihk')
@login_required
def brs_ihk():
    docs = get_docs('IHK', fallback=[])
    return render_template('brs/ihk.html', title='BRS IHK-Inflasi', docs=docs)


@main.route('/brs/ntp')
@login_required
def brs_ntp():
    docs = get_docs('NTP', fallback=[])
    return render_template('brs/ntp.html', title='BRS NTP dan Gabah', docs=docs)


@main.route('/brs/transportasi')
@login_required
def brs_transportasi():
    docs = get_docs('Transportasi', fallback=[])
    return render_template('brs/transportasi.html', title='BRS Transportasi', docs=docs)


@main.route('/brs/ekspor')
@login_required
def brs_ekspor():
    docs = get_docs('Ekspor', fallback=[])
    return render_template('brs/ekspor.html', title='BRS Ekspor Impor', docs=docs)


@main.route('/brs/pariwisata')
@login_required
def brs_pariwisata():
    docs = get_docs('Pariwisata', fallback=[])
    return render_template('brs/pariwisata.html', title='BRS Pariwisata', docs=docs)




# ─── Upload IHK ───────────────────────────────────────────────────────────────

REQUIRED_EXCEL_COLS = [
    'Tahun', 'Bulan', 'Kode Kota', 'Kode Komoditas',
    'IHK', 'NK',
    'Inflasi MtM', 'Inflasi YtD', 'Inflasi YoY',
    'Andil MtM', 'Andil YtD', 'Andil YoY',
]

# ─── Progress Tracker API ─────────────────────────────────────────────────────

@main.route('/upload-progress/<task_id>')
@login_required
def upload_progress(task_id):
    from app.services.sheets import get_progress
    from flask import jsonify
    return jsonify(get_progress(task_id))

@main.route('/upload-ihk', methods=['GET', 'POST'])
@login_required
def upload_ihk():
    if request.method == 'GET':
        return render_template('upload_ihk.html', title='Upload Excel IHK/Inflasi')

    # ── Ambil input form ──────────────────────────────────────────────────────
    bulan  = request.form.get('bulan', '').strip()
    tahun  = request.form.get('tahun', '').strip()
    task_id = request.form.get('task_id', '').strip()
    file   = request.files.get('file')

    # ── Validasi ──────────────────────────────────────────────────────────────
    errors = []
    if not bulan:
        errors.append('Bulan wajib dipilih.')
    if not tahun:
        errors.append('Tahun wajib diisi.')
    elif not tahun.isdigit() or len(tahun) != 4:
        errors.append('Tahun harus 4 digit angka (contoh: 2026).')
    if not file or file.filename == '':
        errors.append('File Excel wajib diunggah.')
    elif not file.filename.lower().endswith('.xlsx'):
        errors.append('File harus berformat .xlsx')

    if errors:
        for e in errors:
            flash(e, 'danger')
        return render_template('upload_ihk.html', title='Upload Excel IHK/Inflasi')

    # ── Baca Excel ────────────────────────────────────────────────────────────
    try:
        import pandas as pd
        # dtype=str: wajib agar kode seperti '0111001' tidak hilang leading zero-nya
        # fillna(''): ganti NaN dengan string kosong
        df = pd.read_excel(file, header=2, dtype=str).fillna('')
    except Exception as e:
        flash(f'Gagal membaca file Excel: {e}', 'danger')
        return render_template('upload_ihk.html', title='Upload Excel IHK/Inflasi')

    # ── Validasi kolom Excel ──────────────────────────────────────────────────
    missing_cols = [c for c in REQUIRED_EXCEL_COLS if c not in df.columns]
    if missing_cols:
        flash(f'Kolom Excel tidak lengkap. Kolom yang kurang: {", ".join(missing_cols)}', 'danger')
        return render_template('upload_ihk.html', title='Upload Excel IHK/Inflasi')

    if df.empty:
        flash('File Excel tidak memiliki data.', 'warning')
        return render_template('upload_ihk.html', title='Upload Excel IHK-Inflasi')

    # ── Validasi & Filter Bulan/Tahun ─────────────────────────────────────────
    # Normalisasi: int(bulan) agar '01' == '1' (kedua sisi)
    bulan_int  = int(bulan)
    tahun_str  = str(tahun).strip()

    # Bandingkan menggunakan int untuk bulan, string untuk tahun
    mask = df.apply(
        lambda r: (
            str(r.get('Tahun', '')).strip() == tahun_str and
            _safe_int(r.get('Bulan', '')) == bulan_int
        ),
        axis=1,
    )

    total_rows  = len(df)
    matched_df  = df[mask]
    matched_cnt = len(matched_df)
    filter_info = None   # pesan tambahan untuk template

    if matched_cnt == 0:
        flash(
            f'Tidak ada data untuk Bulan {bulan} Tahun {tahun} di file Excel. '
            f'Pastikan kolom Bulan dan Tahun di Excel sesuai.',
            'danger'
        )
        return render_template('upload_ihk.html', title='Upload Excel IHK-Inflasi')

    if matched_cnt < total_rows:
        filter_info = (
            f'Hanya {matched_cnt} dari {total_rows} baris yang sesuai '
            f'(Bulan {bulan} / Tahun {tahun}) — baris lain dilewati.'
        )
        df = matched_df   # upload hanya yang cocok

    # ── Generate kode YYMM ────────────────────────────────────────────────────
    col_name = str(tahun)[2:] + str(bulan).zfill(2)   # contoh: 2026 + 08 → '2608'


    # ── Proses ke Google Sheets ───────────────────────────────────────────────
    try:
        from app.services.sheets import process_ihk_upload, clear_progress
        stats = process_ihk_upload(df, col_name, task_id=task_id)

        if stats['inserted'] == 0 and stats['updated'] == 0:
            flash('File berhasil dibaca, namun tidak ada baris data yang valid untuk di-upload.', 'warning')
    except FileNotFoundError:
        flash(
            'File credentials.json tidak ditemukan. '
            'Letakkan service account key di root project.',
            'danger'
        )
        return render_template('upload_ihk.html', title='Upload Excel IHK/Inflasi')
    except Exception as e:
        flash(f'Gagal meng-upload ke Google Sheets: {e}', 'danger')
        return render_template('upload_ihk.html', title='Upload Excel IHK/Inflasi')

    # ── Sukses — langsung render dengan stats (tanpa flash duplikat) ──────────
    if task_id:
        clear_progress(task_id)
        
    return render_template(
        'upload_ihk.html',
        title='Upload Excel IHK-Inflasi',
        stats=stats,
        col_name=col_name,
        filter_info=filter_info,
    )


# ── Download Template Excel ────────────────────────────────────────────────────

@main.route('/download-template-ihk')
@login_required
def download_template_ihk():
    """Generate dan kirim file template Excel IHK-Inflasi."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Template IHK-Inflasi'

    COLS = [
        'Tahun', 'Bulan', 'Kode Kota', 'Nama Kota', 'Kode Komoditas',
        'Nama Komoditas', 'Flag', 'NK', 'IHK',
        'Inflasi MtM', 'Inflasi YtD', 'Inflasi YoY',
        'Andil MtM', 'Andil YtD', 'Andil YoY',
    ]

    # ── Baris 1: Judul ──────────────────────────────────────────────────────
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLS))
    title_cell = ws.cell(row=1, column=1,
        value='Template Upload Data IHK-Inflasi — DIA BRS')
    title_cell.font      = Font(bold=True, size=12, color='FFFFFF')
    title_cell.fill      = PatternFill('solid', fgColor='0D6EFD')
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 24

    # ── Baris 2: Keterangan ──────────────────────────────────────────────────
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(COLS))
    note_cell = ws.cell(row=2, column=1,
        value='Header kolom WAJIB berada di BARIS KE-3. Isi data mulai baris ke-4.')
    note_cell.font      = Font(italic=True, size=10, color='856404')
    note_cell.fill      = PatternFill('solid', fgColor='FFF3CD')
    note_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 18

    # ── Baris 3: Header kolom ───────────────────────────────────────────────
    header_fill   = PatternFill('solid', fgColor='198754')
    header_font   = Font(bold=True, color='FFFFFF', size=10)
    header_border = Border(
        bottom=Side(style='medium', color='145C32'),
        right=Side(style='thin', color='CCCCCC'),
    )
    for col_idx, col_name in enumerate(COLS, start=1):
        cell            = ws.cell(row=3, column=col_idx, value=col_name)
        cell.font       = header_font
        cell.fill       = header_fill
        cell.alignment  = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border     = header_border
    ws.row_dimensions[3].height = 30

    # ── Baris 4: Contoh data ──────────────────────────────────────────────
    sample = [
        2026, 8, 3300, 'PROV JAWA TENGAH', '0',
        'UMUM', 0, 105.12, 105.12,
        0.25, 1.30, 2.10,
        0.05, 0.20, 0.35,
    ]
    sample_fill   = PatternFill('solid', fgColor='E8F5E9')
    sample_font   = Font(color='555555', italic=True, size=10)
    for col_idx, val in enumerate(sample, start=1):
        cell           = ws.cell(row=4, column=col_idx, value=val)
        cell.fill      = sample_fill
        cell.font      = sample_font
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # ── Lebar kolom ──────────────────────────────────────────────────────
    col_widths = [8, 7, 11, 22, 16, 26, 6, 9, 9, 12, 12, 12, 12, 12, 12]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Freeze pada baris 4 (header tetap terlihat saat scroll)
    ws.freeze_panes = 'A4'

    # ── Kirim sebagai response ────────────────────────────────────────────
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    return send_file(
        buf,
        as_attachment=True,
        download_name='template_ihk_inflasi.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


# ─── Upload Ekspor-Impor ───────────────────────────────────────────────────────

REQUIRED_EKSPOR_COLS = [
    'YEAR', 'MTH', 'KODE_HS', 'PEL_MUAT', 'PODAL5',
    'NGRTUJUAN', 'NEWCTRYCOD', 'NETTO', 'FOB', 'ORIG2', 'PROVORIG',
]

REQUIRED_IMPOR_COLS = [
    'HS', 'K_NEGARA', 'N1225',
]

REQUIRED_IMPOR_EXCEL_COLS = [
    'BLN', 'THN_PROSES', 'KODE_HS', 'NEGARA', 'NILAI'
]

REQUIRED_EKSPOR_EXCEL_COLS = [
    'BLN_PROSES', 'THN_PROSES', 'PROVPOD', 'PELABUHAN', 
    'KODE_HS', 'NEGARA', 'NETTO', 'FOB'
]


@main.route('/upload-ekspor-impor', methods=['GET', 'POST'])
@login_required
def upload_ekspor_impor():
    if request.method == 'GET':
        return render_template('upload_ekspor_impor.html', title='Upload Ekspor-Impor')

    form_type = request.form.get('form_type', '').strip()  # 'ekspor' atau 'impor'
    task_id   = request.form.get('task_id', '').strip()
    file      = request.files.get('file')

    # ── Tentukan kolom wajib & label berdasarkan form_type ──────────────────
    if form_type == 'ekspor':
        required_cols = REQUIRED_EKSPOR_COLS
        label         = 'Ekspor'
        active_tab    = 'ekspor'
    elif form_type == 'impor':
        required_cols = REQUIRED_IMPOR_COLS
        label         = 'Impor'
        active_tab    = 'impor'
    else:
        flash('Tipe form tidak valid.', 'danger')
        return render_template('upload_ekspor_impor.html', title='Upload Ekspor-Impor')

    # ── Validasi file ────────────────────────────────────────────────────────
    errors = []
    is_excel = False
    if not file or file.filename == '':
        errors.append('File wajib diunggah.')
    else:
        fn = file.filename.lower()
        if fn.endswith('.dbf'):
            is_excel = False
        elif fn.endswith('.xlsx') or fn.endswith('.xls'):
            is_excel = True
        else:
            errors.append('File harus berformat .dbf atau .xlsx/.xls')

    if errors:
        for e in errors:
            flash(e, 'danger')
        return render_template(
            'upload_ekspor_impor.html',
            title='Upload Ekspor-Impor',
            active_tab=active_tab,
        )

    # ── Baca File (DBF atau Excel) ───────────────────────────────────────────
    import os
    import tempfile
    import pandas as pd
    
    df = None
    temp_path = None
    
    try:
        if is_excel:
            # Baca Excel
            df = pd.read_excel(file, dtype=str).fillna('')
        else:
            # Baca DBF
            from dbfread import DBF
            fd, temp_path = tempfile.mkstemp(suffix='.dbf')
            os.close(fd)
            file.save(temp_path)
            dbf = DBF(temp_path, load=True, lowernames=False)
            df = pd.DataFrame(iter(dbf))
    except Exception as e:
        flash(f'Gagal membaca file: {e}', 'danger')
        return render_template(
            'upload_ekspor_impor.html',
            title='Upload Ekspor-Impor',
            active_tab=active_tab,
        )
    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except:
                pass

    # ── Normalisasi nama kolom (strip whitespace, uppercase) ─────────────────
    df.columns = [c.strip().upper() for c in df.columns]

    # ── Mapping jika Excel Ekspor ───────────────────────────────────────────
    if is_excel and form_type == 'ekspor':
        # Validasi kolom excel khusus ekspor
        missing_excel = [c for c in REQUIRED_EKSPOR_EXCEL_COLS if c not in df.columns]
        if missing_excel:
            flash(
                f'Kolom Excel Ekspor tidak lengkap. Kolom yang kurang: {", ".join(missing_excel)}',
                'danger',
            )
            return render_template('upload_ekspor_impor.html', title='Upload Ekspor-Impor', active_tab=active_tab)
        
        # Rename ke format yang diharapkan process_ekspor_upload
        mapping = {
            'THN_PROSES': 'YEAR',
            'BLN_PROSES': 'MTH',
            'KODE_HS': 'KODE_HS',
            'PELABUHAN': 'PODAL5',
            'NEGARA': 'NEWCTRYCOD',
            'NETTO': 'NETTO',
            'FOB': 'FOB'
        }
        df = df.rename(columns=mapping)
        # Tambahkan kolom dummy yang diabaikan tapi mungkin dicek (required_cols)
        if 'PEL_MUAT' not in df.columns: df['PEL_MUAT'] = ''
        if 'NGRTUJUAN' not in df.columns: df['NGRTUJUAN'] = ''
        if 'ORIG2' not in df.columns: df['ORIG2'] = ''
        if 'PROVORIG' not in df.columns: df['PROVORIG'] = ''
        
        required_cols = REQUIRED_EKSPOR_COLS # Gunakan standard required cols setelah rename

    # ── Mapping jika Excel/DBF Impor ──────────────────────────────────────────
    if form_type == 'impor':
        if is_excel:
            missing_excel = [c for c in REQUIRED_IMPOR_EXCEL_COLS if c not in df.columns]
            if missing_excel:
                flash(
                    f'Kolom Excel Impor tidak lengkap. Kurang: {", ".join(missing_excel)}',
                    'danger',
                )
                return render_template('upload_ekspor_impor.html', title='Upload Ekspor-Impor', active_tab=active_tab)
            
            # Rename Excel Impor -> Standard DBF Impor format
            mapping_impor = {
                'KODE_HS': 'HS',
                'NEGARA': 'K_NEGARA',
                'NILAI': 'N1225',
                'THN_PROSES': 'YEAR',
                'BLN': 'MTH'
            }
            df = df.rename(columns=mapping_impor)
        else:
            # Jika DBF Impor, pastikan form Bulan dan Tahun terisi
            bulan = request.form.get('bulan', '').strip()
            tahun = request.form.get('tahun', '').strip()
            
            if not bulan or not tahun:
                flash('Bulan dan Tahun wajib diisi untuk upload DBF Impor.', 'danger')
                return render_template('upload_ekspor_impor.html', title='Upload Ekspor-Impor', active_tab=active_tab)
            
            # Tambahkan ke dataframe agar strukturnya sama dengan ekspor / excel impor
            df['YEAR'] = str(tahun)
            df['MTH'] = str(bulan).zfill(2)

        required_cols = REQUIRED_IMPOR_COLS + ['YEAR', 'MTH']

    # ── Validasi kolom (General) ──────────────────────────────────────────────
    missing_cols = [c for c in required_cols if c not in df.columns]
    if missing_cols:
        flash(
            f'Kolom {label} tidak lengkap. Kolom yang kurang: {", ".join(missing_cols)}',
            'danger',
        )
        return render_template(
            'upload_ekspor_impor.html',
            title='Upload Ekspor-Impor',
            active_tab=active_tab,
        )

    if df.empty:
        flash(f'File {label} tidak memiliki data.', 'warning')
        return render_template(
            'upload_ekspor_impor.html',
            title='Upload Ekspor-Impor',
            active_tab=active_tab,
        )

    row_count = len(df)
    
    stats = None
    try:
        from app.services.sheets import clear_progress
        if form_type == 'ekspor':
            from app.services.sheets import process_ekspor_upload
            stats = process_ekspor_upload(df, task_id=task_id)
        elif form_type == 'impor':
            from app.services.sheets import process_impor_upload
            stats = process_impor_upload(df, task_id=task_id)
    except Exception as e:
        flash(f'Gagal meng-upload {label} ke Google Sheets: {e}', 'danger')
        return render_template(
            'upload_ekspor_impor.html',
            title='Upload Ekspor-Impor',
            active_tab=active_tab,
        )
        
    if task_id:
        clear_progress(task_id)

    return render_template(
        'upload_ekspor_impor.html',
        title='Upload Ekspor-Impor',
        active_tab=active_tab,
        last_upload={'type': label, 'rows': row_count},
        stats=stats
    )


# ── Download Template Ekspor ───────────────────────────────────────────────────

@main.route('/download-template-ekspor')
@login_required
def download_template_ekspor():
    """Generate dan kirim file template Excel Ekspor."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    COLS = REQUIRED_EKSPOR_COLS
    SAMPLE = ['2025', '1', '0101', 'TANJUNG MAS', 'JT', 'USA', 'US', '1000.50', '5000.00', 'ID', 'JAWA TENGAH']

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Template Ekspor'

    # Baris 1: Judul
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLS))
    title_cell = ws.cell(row=1, column=1, value='Template Upload Data Ekspor — DIA BRS')
    title_cell.font      = Font(bold=True, size=12, color='FFFFFF')
    title_cell.fill      = PatternFill('solid', fgColor='0D9488')
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 24

    # Baris 2: Keterangan
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(COLS))
    note_cell = ws.cell(row=2, column=1, value='Header kolom WAJIB berada di BARIS KE-1. Isi data mulai baris ke-2.')
    note_cell.font      = Font(italic=True, size=10, color='856404')
    note_cell.fill      = PatternFill('solid', fgColor='FFF3CD')
    note_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 18

    # Baris 3: Header
    hfill  = PatternFill('solid', fgColor='0D9488')
    hfont  = Font(bold=True, color='FFFFFF', size=10)
    hbord  = Border(bottom=Side(style='medium', color='065F46'), right=Side(style='thin', color='CCCCCC'))
    for ci, col in enumerate(COLS, start=1):
        cell            = ws.cell(row=3, column=ci, value=col)
        cell.font       = hfont
        cell.fill       = hfill
        cell.alignment  = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border     = hbord
    ws.row_dimensions[3].height = 30

    # Baris 4: Contoh data
    sfill = PatternFill('solid', fgColor='ECFDF5')
    sfont = Font(color='555555', italic=True, size=10)
    for ci, val in enumerate(SAMPLE, start=1):
        cell           = ws.cell(row=4, column=ci, value=val)
        cell.fill      = sfill
        cell.font      = sfont
        cell.alignment = Alignment(horizontal='center', vertical='center')

    col_widths = [8, 6, 10, 16, 10, 14, 14, 12, 12, 8, 16]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A4'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name='template_ekspor.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


# ── Download Template Impor ────────────────────────────────────────────────────

@main.route('/download-template-impor')
@login_required
def download_template_impor():
    """Generate dan kirim file template Excel Impor."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    COLS   = REQUIRED_IMPOR_EXCEL_COLS
    SAMPLE = ['1', '2025', 'TANJUNG MAS', '0101.20.10', 'UNITED STATES', '450.00']

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Template Impor'

    # Baris 1: Judul
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLS))
    title_cell = ws.cell(row=1, column=1, value='Template Upload Data Impor — DIA BRS')
    title_cell.font      = Font(bold=True, size=12, color='FFFFFF')
    title_cell.fill      = PatternFill('solid', fgColor='1D4ED8')
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 24

    # Baris 2: Keterangan
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(COLS))
    note_cell = ws.cell(row=2, column=1, value='Header kolom WAJIB berada di BARIS KE-1. Isi data mulai baris ke-2.')
    note_cell.font      = Font(italic=True, size=10, color='856404')
    note_cell.fill      = PatternFill('solid', fgColor='FFF3CD')
    note_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 18

    # Baris 3: Header
    hfill  = PatternFill('solid', fgColor='1D4ED8')
    hfont  = Font(bold=True, color='FFFFFF', size=10)
    hbord  = Border(bottom=Side(style='medium', color='1E3A8A'), right=Side(style='thin', color='CCCCCC'))
    for ci, col in enumerate(COLS, start=1):
        cell            = ws.cell(row=3, column=ci, value=col)
        cell.font       = hfont
        cell.fill       = hfill
        cell.alignment  = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border     = hbord
    ws.row_dimensions[3].height = 30

    # Baris 4: Contoh data
    sfill = PatternFill('solid', fgColor='EFF6FF')
    sfont = Font(color='555555', italic=True, size=10)
    for ci, val in enumerate(SAMPLE, start=1):
        cell           = ws.cell(row=4, column=ci, value=val)
        cell.fill      = sfill
        cell.font      = sfont
        cell.alignment = Alignment(horizontal='center', vertical='center')

    col_widths = [8, 12, 16, 14, 16, 14]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A4'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name='template_impor.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


# ─── Upload NTP ───────────────────────────────────────────────────────────────

REQUIRED_NTP_COLS = [
    'Tahun', 'Bulan', 'Kode Provinsi', 'Nama Provinsi',
    'NTP', 'NTPP', 'NTPH', 'NTPN', 'NTPF', 'NTPE',
]

@main.route('/upload-ntp', methods=['GET', 'POST'])
@login_required
def upload_ntp():
    if request.method == 'GET':
        return render_template('upload_ntp.html', title='Upload Excel NTP dan Gabah')

    bulan   = request.form.get('bulan', '').strip()
    tahun   = request.form.get('tahun', '').strip()
    file    = request.files.get('file')

    errors = []
    if not bulan:
        errors.append('Bulan wajib dipilih.')
    if not tahun:
        errors.append('Tahun wajib diisi.')
    elif not tahun.isdigit() or len(tahun) != 4:
        errors.append('Tahun harus 4 digit angka.')
    if not file or file.filename == '':
        errors.append('File Excel wajib diunggah.')
    elif not file.filename.lower().endswith('.xlsx'):
        errors.append('File harus berformat .xlsx')

    if errors:
        for e in errors:
            flash(e, 'danger')
        return render_template('upload_ntp.html', title='Upload Excel NTP dan Gabah')

    try:
        import pandas as pd
        df = pd.read_excel(file, header=2, dtype=str).fillna('')
    except Exception as e:
        flash(f'Gagal membaca file Excel: {e}', 'danger')
        return render_template('upload_ntp.html', title='Upload Excel NTP dan Gabah')

    missing_cols = [c for c in REQUIRED_NTP_COLS if c not in df.columns]
    if missing_cols:
        flash(f'Kolom Excel tidak lengkap. Kolom yang kurang: {", ".join(missing_cols)}', 'danger')
        return render_template('upload_ntp.html', title='Upload Excel NTP dan Gabah')

    if df.empty:
        flash('File Excel tidak memiliki data.', 'warning')
        return render_template('upload_ntp.html', title='Upload Excel NTP dan Gabah')

    flash('File berhasil dibaca. Fitur upload NTP ke Google Sheets sedang dalam pengembangan.', 'info')
    return render_template('upload_ntp.html', title='Upload Excel NTP dan Gabah')


@main.route('/download-template-ntp')
@login_required
def download_template_ntp():
    """Generate dan kirim file template Excel NTP dan Gabah."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    COLS = REQUIRED_NTP_COLS
    SAMPLE = ['2026', '8', '33', 'JAWA TENGAH', '105.12', '102.30', '107.50', '103.00', '108.20', '104.10']

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Template NTP'

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLS))
    c = ws.cell(row=1, column=1, value='Template Upload Data NTP dan Gabah — DIA BRS')
    c.font = Font(bold=True, size=12, color='FFFFFF')
    c.fill = PatternFill('solid', fgColor='198754')
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 24

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(COLS))
    n = ws.cell(row=2, column=1, value='Header kolom WAJIB berada di BARIS KE-3. Isi data mulai baris ke-4.')
    n.font = Font(italic=True, size=10, color='856404')
    n.fill = PatternFill('solid', fgColor='FFF3CD')
    n.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 18

    hfill = PatternFill('solid', fgColor='198754')
    hfont = Font(bold=True, color='FFFFFF', size=10)
    hbord = Border(bottom=Side(style='medium', color='145C32'), right=Side(style='thin', color='CCCCCC'))
    for ci, col in enumerate(COLS, start=1):
        cell = ws.cell(row=3, column=ci, value=col)
        cell.font, cell.fill = hfont, hfill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = hbord
    ws.row_dimensions[3].height = 30

    sfill = PatternFill('solid', fgColor='E8F5E9')
    sfont = Font(color='555555', italic=True, size=10)
    for ci, val in enumerate(SAMPLE[:len(COLS)], start=1):
        cell = ws.cell(row=4, column=ci, value=val)
        cell.fill, cell.font = sfill, sfont
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for i in range(1, len(COLS) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 14
    ws.freeze_panes = 'A4'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name='template_ntp.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ─── Upload Pariwisata ────────────────────────────────────────────────────────

REQUIRED_PARIWISATA_COLS = [
    'Tahun', 'Bulan', 'Nama Hotel', 'Kelas Hotel',
    'TPK', 'RATS', 'Tamu Asing', 'Tamu Domestik',
]

@main.route('/upload-pariwisata', methods=['GET', 'POST'])
@login_required
def upload_pariwisata():
    if request.method == 'GET':
        return render_template('upload_pariwisata.html', title='Upload Excel Pariwisata')

    bulan   = request.form.get('bulan', '').strip()
    tahun   = request.form.get('tahun', '').strip()
    file    = request.files.get('file')

    errors = []
    if not bulan:
        errors.append('Bulan wajib dipilih.')
    if not tahun:
        errors.append('Tahun wajib diisi.')
    elif not tahun.isdigit() or len(tahun) != 4:
        errors.append('Tahun harus 4 digit angka.')
    if not file or file.filename == '':
        errors.append('File Excel wajib diunggah.')
    elif not file.filename.lower().endswith('.xlsx'):
        errors.append('File harus berformat .xlsx')

    if errors:
        for e in errors:
            flash(e, 'danger')
        return render_template('upload_pariwisata.html', title='Upload Excel Pariwisata')

    try:
        import pandas as pd
        df = pd.read_excel(file, header=2, dtype=str).fillna('')
    except Exception as e:
        flash(f'Gagal membaca file Excel: {e}', 'danger')
        return render_template('upload_pariwisata.html', title='Upload Excel Pariwisata')

    missing_cols = [c for c in REQUIRED_PARIWISATA_COLS if c not in df.columns]
    if missing_cols:
        flash(f'Kolom Excel tidak lengkap. Kolom yang kurang: {", ".join(missing_cols)}', 'danger')
        return render_template('upload_pariwisata.html', title='Upload Excel Pariwisata')

    if df.empty:
        flash('File Excel tidak memiliki data.', 'warning')
        return render_template('upload_pariwisata.html', title='Upload Excel Pariwisata')

    flash('File berhasil dibaca. Fitur upload Pariwisata ke Google Sheets sedang dalam pengembangan.', 'info')
    return render_template('upload_pariwisata.html', title='Upload Excel Pariwisata')


@main.route('/download-template-pariwisata')
@login_required
def download_template_pariwisata():
    """Generate dan kirim file template Excel Pariwisata."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    COLS = REQUIRED_PARIWISATA_COLS
    SAMPLE = ['2026', '8', 'Grand Artos', 'Bintang 4', '62.50', '2.10', '120', '850']

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Template Pariwisata'

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLS))
    c = ws.cell(row=1, column=1, value='Template Upload Data Pariwisata — DIA BRS')
    c.font = Font(bold=True, size=12, color='FFFFFF')
    c.fill = PatternFill('solid', fgColor='7C3AED')
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 24

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(COLS))
    n = ws.cell(row=2, column=1, value='Header kolom WAJIB berada di BARIS KE-3. Isi data mulai baris ke-4.')
    n.font = Font(italic=True, size=10, color='856404')
    n.fill = PatternFill('solid', fgColor='FFF3CD')
    n.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 18

    hfill = PatternFill('solid', fgColor='7C3AED')
    hfont = Font(bold=True, color='FFFFFF', size=10)
    hbord = Border(bottom=Side(style='medium', color='5B21B6'), right=Side(style='thin', color='CCCCCC'))
    for ci, col in enumerate(COLS, start=1):
        cell = ws.cell(row=3, column=ci, value=col)
        cell.font, cell.fill = hfont, hfill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = hbord
    ws.row_dimensions[3].height = 30

    sfill = PatternFill('solid', fgColor='EDE9FE')
    sfont = Font(color='555555', italic=True, size=10)
    for ci, val in enumerate(SAMPLE[:len(COLS)], start=1):
        cell = ws.cell(row=4, column=ci, value=val)
        cell.fill, cell.font = sfill, sfont
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for i in range(1, len(COLS) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 16
    ws.freeze_panes = 'A4'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name='template_pariwisata.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ─── Upload Transportasi ──────────────────────────────────────────────────────

REQUIRED_TRANSPORTASI_COLS = [
    'Tahun', 'Bulan', 'Moda', 'Nama Perusahaan',
    'Penumpang Berangkat', 'Penumpang Datang', 'Barang Muat', 'Barang Bongkar',
]

@main.route('/upload-transportasi', methods=['GET', 'POST'])
@login_required
def upload_transportasi():
    if request.method == 'GET':
        return render_template('upload_transportasi.html', title='Upload Excel Transportasi')

    bulan   = request.form.get('bulan', '').strip()
    tahun   = request.form.get('tahun', '').strip()
    file    = request.files.get('file')

    errors = []
    if not bulan:
        errors.append('Bulan wajib dipilih.')
    if not tahun:
        errors.append('Tahun wajib diisi.')
    elif not tahun.isdigit() or len(tahun) != 4:
        errors.append('Tahun harus 4 digit angka.')
    if not file or file.filename == '':
        errors.append('File Excel wajib diunggah.')
    elif not file.filename.lower().endswith('.xlsx'):
        errors.append('File harus berformat .xlsx')

    if errors:
        for e in errors:
            flash(e, 'danger')
        return render_template('upload_transportasi.html', title='Upload Excel Transportasi')

    try:
        import pandas as pd
        df = pd.read_excel(file, header=2, dtype=str).fillna('')
    except Exception as e:
        flash(f'Gagal membaca file Excel: {e}', 'danger')
        return render_template('upload_transportasi.html', title='Upload Excel Transportasi')

    missing_cols = [c for c in REQUIRED_TRANSPORTASI_COLS if c not in df.columns]
    if missing_cols:
        flash(f'Kolom Excel tidak lengkap. Kolom yang kurang: {", ".join(missing_cols)}', 'danger')
        return render_template('upload_transportasi.html', title='Upload Excel Transportasi')

    if df.empty:
        flash('File Excel tidak memiliki data.', 'warning')
        return render_template('upload_transportasi.html', title='Upload Excel Transportasi')

    flash('File berhasil dibaca. Fitur upload Transportasi ke Google Sheets sedang dalam pengembangan.', 'info')
    return render_template('upload_transportasi.html', title='Upload Excel Transportasi')


@main.route('/download-template-transportasi')
@login_required
def download_template_transportasi():
    """Generate dan kirim file template Excel Transportasi."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    COLS = REQUIRED_TRANSPORTASI_COLS
    SAMPLE = ['2026', '8', 'Udara', 'Garuda Indonesia', '12500', '11800', '850.50', '920.30']

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Template Transportasi'

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLS))
    c = ws.cell(row=1, column=1, value='Template Upload Data Transportasi — DIA BRS')
    c.font = Font(bold=True, size=12, color='FFFFFF')
    c.fill = PatternFill('solid', fgColor='D97706')
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 24

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(COLS))
    n = ws.cell(row=2, column=1, value='Header kolom WAJIB berada di BARIS KE-3. Isi data mulai baris ke-4.')
    n.font = Font(italic=True, size=10, color='856404')
    n.fill = PatternFill('solid', fgColor='FFF3CD')
    n.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 18

    hfill = PatternFill('solid', fgColor='D97706')
    hfont = Font(bold=True, color='FFFFFF', size=10)
    hbord = Border(bottom=Side(style='medium', color='92400E'), right=Side(style='thin', color='CCCCCC'))
    for ci, col in enumerate(COLS, start=1):
        cell = ws.cell(row=3, column=ci, value=col)
        cell.font, cell.fill = hfont, hfill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = hbord
    ws.row_dimensions[3].height = 30

    sfill = PatternFill('solid', fgColor='FFFBEB')
    sfont = Font(color='555555', italic=True, size=10)
    for ci, val in enumerate(SAMPLE[:len(COLS)], start=1):
        cell = ws.cell(row=4, column=ci, value=val)
        cell.fill, cell.font = sfill, sfont
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for i in range(1, len(COLS) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 18
    ws.freeze_panes = 'A4'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name='template_transportasi.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
