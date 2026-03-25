"""
Routes untuk blueprint main (halaman utama dan BRS).
"""
from functools import wraps
from flask import render_template, request, redirect, url_for, flash, send_file, abort
from flask_login import login_required, current_user
from app.main import main
from app.services.sheets import get_docs, get_klik_links, clear_docs_cache
import app.services.sheets as _sheets
import io

from app.main.brs_cols import BRS_CONFIG

@main.context_processor
def inject_brs_config():
    return dict(brs_config=BRS_CONFIG)

def _require_akses(flag_key):
    """Decorator: tolak akses (403) jika user tidak punya flag_key."""
    def decorator(f):
        @wraps(f)
        def wrapped(*args, **kwargs):
            if not current_user.akses.get(flag_key, False):
                abort(403)
            return f(*args, **kwargs)
        return wrapped
    return decorator



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


@main.route('/dia-brs/klik-refresh')
def klik_refresh():
    """Hapus cache Klik Distribusi lalu redirect ke halaman utama."""
    from app.services.sheets import _klik_cache, _klik_cache_ts
    import app.services.sheets as _sheets
    _sheets._klik_cache    = []
    _sheets._klik_cache_ts = 0.0
    flash('Klik Distribusi telah di-refresh — Daftar link berhasil diperbarui.', 'success')
    return redirect(url_for('main.index'))


@main.route('/dia-brs/docs-refresh')
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


@main.route('/dia-brs/dashboard')
@login_required
def dashboard():
    return redirect(url_for('main.dia_brs'))


@main.route('/dia-brs')
@login_required
def dia_brs():
    return render_template('dashboard.html', title='DIA BRS — Dashboard')


from app.services.sheets import get_docs


@main.route('/dia-brs/ihk')
@login_required
@_require_akses('akses_ihk')
def brs_ihk():
    docs = get_docs('IHK', fallback=[])
    return render_template('brs/ihk.html', title='BRS IHK-Inflasi', docs=docs)


@main.route('/dia-brs/ntp')
@login_required
@_require_akses('akses_ntp')
def brs_ntp():
    docs = get_docs('NTP', fallback=[])
    return render_template('brs/ntp.html', title='BRS NTP', docs=docs)


@main.route('/dia-brs/transportasi')
@login_required
@_require_akses('akses_transportasi')
def brs_transportasi():
    docs = get_docs('Transportasi', fallback=[])
    return render_template('brs/transportasi.html', title='BRS Transportasi', docs=docs)


@main.route('/dia-brs/ekspor-impor')
@login_required
@_require_akses('akses_ekspor_impor')
def brs_ekspor():
    docs = get_docs('Ekspor', fallback=[])
    return render_template('brs/ekspor.html', title='BRS Ekspor Impor', docs=docs)


@main.route('/dia-brs/pariwisata')
@login_required
@_require_akses('akses_pariwisata')
def brs_pariwisata():
    docs = get_docs('Pariwisata', fallback=[])
    return render_template('brs/pariwisata.html', title='BRS Pariwisata', docs=docs)




# ─── Upload IHK ───────────────────────────────────────────────────────────────

# ─── Progress Tracker API ─────────────────────────────────────────────────────

@main.route('/dia-brs/upload-progress/<task_id>')
@login_required
def upload_progress(task_id):
    from app.services.sheets import get_progress
    from flask import jsonify
    return jsonify(get_progress(task_id))

@main.route('/dia-brs/ihk/upload', methods=['GET', 'POST'])
@login_required
@_require_akses('akses_ihk')
def upload_ihk():
    if request.method == 'GET':
        return render_template('brs/upload_ihk.html', title='Upload Excel IHK/Inflasi')

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
        return render_template('brs/upload_ihk.html', title='Upload Excel IHK/Inflasi')

    # ── Baca Excel ────────────────────────────────────────────────────────────
    try:
        import pandas as pd
        # dtype=str: wajib agar kode seperti '0111001' tidak hilang leading zero-nya
        # fillna(''): ganti NaN dengan string kosong
        df = pd.read_excel(file, header=2, dtype=str).fillna('')
    except Exception as e:
        flash(f'Gagal membaca file Excel: {e}', 'danger')
        return render_template('brs/upload_ihk.html', title='Upload Excel IHK/Inflasi')

    # ── Validasi kolom Excel ──────────────────────────────────────────────────
    missing_cols = [c for c in BRS_CONFIG['ihk']['required_cols'] if c not in df.columns]
    if missing_cols:
        flash(f'Kolom Excel tidak lengkap. Kolom yang kurang: {", ".join(missing_cols)}', 'danger')
        return render_template('brs/upload_ihk.html', title='Upload Excel IHK/Inflasi')

    if df.empty:
        flash('File Excel tidak memiliki data.', 'warning')
        return render_template('brs/upload_ihk.html', title='Upload Excel IHK-Inflasi')

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
        return render_template('brs/upload_ihk.html', title='Upload Excel IHK-Inflasi')

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
        return render_template('brs/upload_ihk.html', title='Upload Excel IHK/Inflasi')
    except Exception as e:
        flash(f'Gagal meng-upload ke Google Sheets: {e}', 'danger')
        return render_template('brs/upload_ihk.html', title='Upload Excel IHK/Inflasi')

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

@main.route('/dia-brs/ihk/template')
@login_required
@_require_akses('akses_ihk')
def download_template_ihk():
    """Generate dan kirim file template Excel IHK-Inflasi."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Template IHK-Inflasi'

    COLS = BRS_CONFIG['ihk']['required_cols']

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
    sample = BRS_CONFIG['ihk']['template']['sample']
    # sample_fill   = PatternFill('solid', fgColor='E8F5E9')
    # sample_font   = Font(color='555555', italic=True, size=10)
    for col_idx, val in enumerate(sample, start=1):
        cell           = ws.cell(row=4, column=col_idx, value=val)
        # cell.fill      = sample_fill
        # cell.font      = sample_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        if isinstance(val, float):
            cell.number_format = '#,##0.00'

    # ── Lebar kolom ──────────────────────────────────────────────────────
    col_widths = BRS_CONFIG['ihk']['template']['col_widths']
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


@main.route('/dia-brs/ekspor-impor/upload', methods=['GET', 'POST'])
@login_required
@_require_akses('akses_ekspor_impor')
def upload_ekspor_impor():
    if request.method == 'GET':
        return render_template('brs/upload_ekspor_impor.html', title='Upload Ekspor-Impor')

    form_type = request.form.get('form_type', '').strip()  # 'ekspor' atau 'impor'
    task_id   = request.form.get('task_id', '').strip()
    file      = request.files.get('file')

    # ── Tentukan kolom wajib & label berdasarkan form_type ──────────────────
    if form_type == 'ekspor':
        required_cols = BRS_CONFIG['ekspor']['required_cols']
        label         = 'Ekspor'
        active_tab    = 'ekspor'
    elif form_type == 'impor':
        required_cols = BRS_CONFIG['impor']['required_cols']
        label         = 'Impor'
        active_tab    = 'impor'
    else:
        flash('Tipe form tidak valid.', 'danger')
        return render_template('brs/upload_ekspor_impor.html', title='Upload Ekspor-Impor')

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
        missing_excel = [c for c in BRS_CONFIG['ekspor']['excel_cols'] if c not in df.columns]
        if missing_excel:
            flash(
                f'Kolom Excel Ekspor tidak lengkap. Kolom yang kurang: {", ".join(missing_excel)}',
                'danger',
            )
            return render_template('brs/upload_ekspor_impor.html', title='Upload Ekspor-Impor', active_tab=active_tab)
        
        # Rename ke format yang diharapkan process_ekspor_upload
        mapping = {
            'THN_PROSES': 'YEAR',
            'BLN_PROSES': 'MTH',
            'KODE_HS': 'KODE_HS',
            'PELABUHAN': 'PODAL5',
            'NEGARA': 'NEWCTRYCOD',
            'FOB': 'FOB'
        }
        df = df.rename(columns=mapping)
        
        required_cols = BRS_CONFIG['ekspor']['required_cols'] # Gunakan standard required cols setelah rename

    # ── Mapping jika Excel/DBF Impor ──────────────────────────────────────────
    if form_type == 'impor':
        if is_excel:
            missing_excel = [c for c in BRS_CONFIG['impor']['excel_cols'] if c not in df.columns]
            if missing_excel:
                flash(
                    f'Kolom Excel Impor tidak lengkap. Kurang: {", ".join(missing_excel)}',
                    'danger',
                )
                return render_template('brs/upload_ekspor_impor.html', title='Upload Ekspor-Impor', active_tab=active_tab)
            
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
                return render_template('brs/upload_ekspor_impor.html', title='Upload Ekspor-Impor', active_tab=active_tab)
            
            # Tambahkan ke dataframe agar strukturnya sama dengan ekspor / excel impor
            df['YEAR'] = str(tahun)
            df['MTH'] = str(bulan).zfill(2)
            
            # Dinamis NMMYY
            expected_n_col = f"N{str(bulan).zfill(2)}{str(tahun)[-2:]}"
            
            if expected_n_col not in df.columns:
                flash(f'Kolom DBF Impor tidak lengkap. Kolom yang kurang: {expected_n_col}', 'danger')
                return render_template('brs/upload_ekspor_impor.html', title='Upload Ekspor-Impor', active_tab=active_tab)
                
            df = df.rename(columns={expected_n_col: 'N1225'})

        required_cols = BRS_CONFIG['impor']['required_cols'] + ['N1225', 'YEAR', 'MTH']

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

@main.route('/dia-brs/ekspor-impor/template-ekspor')
@login_required
@_require_akses('akses_ekspor_impor')
def download_template_ekspor():
    """Generate dan kirim file template Excel Ekspor."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    COLS = BRS_CONFIG['ekspor']['excel_cols']
    SAMPLE = BRS_CONFIG['ekspor']['template']['sample']

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Template Ekspor'

    # Baris 1: Header
    hfill  = PatternFill('solid', fgColor='0D9488')
    hfont  = Font(bold=True, color='FFFFFF', size=10)
    hbord  = Border(bottom=Side(style='medium', color='065F46'), right=Side(style='thin', color='CCCCCC'))
    for ci, col in enumerate(COLS, start=1):
        cell            = ws.cell(row=1, column=ci, value=col)
        cell.font       = hfont
        cell.fill       = hfill
        cell.alignment  = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border     = hbord
    ws.row_dimensions[1].height = 30

    # Baris 2: Contoh data
    # sfill = PatternFill('solid', fgColor='ECFDF5')
    # sfont = Font(color='555555', italic=True, size=10)
    for ci, val in enumerate(SAMPLE, start=1):
        cell           = ws.cell(row=2, column=ci, value=val)
        # cell.fill      = sfill
        # cell.font      = sfont
        cell.alignment = Alignment(horizontal='center', vertical='center')

    col_widths = BRS_CONFIG['ekspor']['template']['col_widths']
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'

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

@main.route('/dia-brs/ekspor-impor/template-impor')
@login_required
@_require_akses('akses_ekspor_impor')
def download_template_impor():
    """Generate dan kirim file template Excel Impor."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    COLS   = BRS_CONFIG['impor']['excel_cols']
    SAMPLE = BRS_CONFIG['impor']['template']['sample']

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Template Impor'

    # Baris 1: Header
    hfill  = PatternFill('solid', fgColor='1D4ED8')
    hfont  = Font(bold=True, color='FFFFFF', size=10)
    hbord  = Border(bottom=Side(style='medium', color='1E3A8A'), right=Side(style='thin', color='CCCCCC'))
    for ci, col in enumerate(COLS, start=1):
        cell            = ws.cell(row=1, column=ci, value=col)
        cell.font       = hfont
        cell.fill       = hfill
        cell.alignment  = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border     = hbord
    ws.row_dimensions[1].height = 30

    # Baris 2: Contoh data
    # sfill = PatternFill('solid', fgColor='EFF6FF')
    # sfont = Font(color='555555', italic=True, size=10)
    for ci, val in enumerate(SAMPLE, start=1):
        cell           = ws.cell(row=2, column=ci, value=val)
        # cell.fill      = sfill
        # cell.font      = sfont
        cell.alignment = Alignment(horizontal='center', vertical='center')

    col_widths = BRS_CONFIG['impor']['template']['col_widths']
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'

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

@main.route('/dia-brs/ntp/upload', methods=['GET', 'POST'])
@login_required
@_require_akses('akses_ntp')
def upload_ntp():
    if request.method == 'GET':
        return render_template('brs/upload_ntp.html', title='Upload Excel NTP')

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
        return render_template('brs/upload_ntp.html', title='Upload Excel NTP')

    try:
        import pandas as pd
        df = pd.read_excel(file, header=2, dtype=str).fillna('')
    except Exception as e:
        flash(f'Gagal membaca file Excel: {e}', 'danger')
        return render_template('brs/upload_ntp.html', title='Upload Excel NTP')

    missing_cols = [c for c in BRS_CONFIG['ntp']['required_cols'] if c not in df.columns]
    if missing_cols:
        flash(f'Kolom Excel tidak lengkap. Kolom yang kurang: {", ".join(missing_cols)}', 'danger')
        return render_template('brs/upload_ntp.html', title='Upload Excel NTP')

    if df.empty:
        flash('File Excel tidak memiliki data.', 'warning')
        return render_template('brs/upload_ntp.html', title='Upload Excel NTP')

    flash('File berhasil dibaca. Fitur upload NTP ke Google Sheets sedang dalam pengembangan.', 'info')
    return render_template('brs/upload_ntp.html', title='Upload Excel NTP')


@main.route('/dia-brs/ntp/template')
@login_required
@_require_akses('akses_ntp')
def download_template_ntp():
    """Generate dan kirim file template Excel NTP."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    COLS = BRS_CONFIG['ntp']['required_cols']
    SAMPLE = BRS_CONFIG['ntp']['template']['sample']

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Template NTP'

    ws.merge_cells('A1:J1')
    c = ws.cell(row=1, column=1, value='Template Upload Data NTP — DIA BRS')
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

@main.route('/dia-brs/pariwisata/upload', methods=['GET', 'POST'])
@login_required
@_require_akses('akses_pariwisata')
def upload_pariwisata():
    if request.method == 'GET':
        return render_template('brs/upload_pariwisata.html', title='Upload Excel Pariwisata')

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
        return render_template('brs/upload_pariwisata.html', title='Upload Excel Pariwisata')

    try:
        import pandas as pd
        df = pd.read_excel(file, header=2, dtype=str).fillna('')
    except Exception as e:
        flash(f'Gagal membaca file Excel: {e}', 'danger')
        return render_template('brs/upload_pariwisata.html', title='Upload Excel Pariwisata')

    missing_cols = [c for c in BRS_CONFIG['pariwisata']['required_cols'] if c not in df.columns]
    if missing_cols:
        flash(f'Kolom Excel tidak lengkap. Kolom yang kurang: {", ".join(missing_cols)}', 'danger')
        return render_template('brs/upload_pariwisata.html', title='Upload Excel Pariwisata')

    if df.empty:
        flash('File Excel tidak memiliki data.', 'warning')
        return render_template('brs/upload_pariwisata.html', title='Upload Excel Pariwisata')

    flash('File berhasil dibaca. Fitur upload Pariwisata ke Google Sheets sedang dalam pengembangan.', 'info')
    return render_template('brs/upload_pariwisata.html', title='Upload Excel Pariwisata')


@main.route('/dia-brs/pariwisata/template')
@login_required
@_require_akses('akses_pariwisata')
def download_template_pariwisata():
    """Generate dan kirim file template Excel Pariwisata."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    COLS = BRS_CONFIG['pariwisata']['required_cols']
    SAMPLE = BRS_CONFIG['pariwisata']['template']['sample']

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

@main.route('/dia-brs/transportasi/upload', methods=['GET', 'POST'])
@login_required
@_require_akses('akses_transportasi')
def upload_transportasi():
    if request.method == 'GET':
        return render_template('brs/upload_transportasi.html', title='Upload Excel Transportasi')

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
        return render_template('brs/upload_transportasi.html', title='Upload Excel Transportasi')

    try:
        import pandas as pd
        df = pd.read_excel(file, header=2, dtype=str).fillna('')
    except Exception as e:
        flash(f'Gagal membaca file Excel: {e}', 'danger')
        return render_template('brs/upload_transportasi.html', title='Upload Excel Transportasi')

    missing_cols = [c for c in BRS_CONFIG['transportasi']['required_cols'] if c not in df.columns]
    if missing_cols:
        flash(f'Kolom Excel tidak lengkap. Kolom yang kurang: {", ".join(missing_cols)}', 'danger')
        return render_template('brs/upload_transportasi.html', title='Upload Excel Transportasi')

    if df.empty:
        flash('File Excel tidak memiliki data.', 'warning')
        return render_template('brs/upload_transportasi.html', title='Upload Excel Transportasi')

    flash('File berhasil dibaca. Fitur upload Transportasi ke Google Sheets sedang dalam pengembangan.', 'info')
    return render_template('brs/upload_transportasi.html', title='Upload Excel Transportasi')


@main.route('/dia-brs/transportasi/template')
@login_required
@_require_akses('akses_transportasi')
def download_template_transportasi():
    """Generate dan kirim file template Excel Transportasi."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    COLS = BRS_CONFIG['transportasi']['required_cols']
    SAMPLE = BRS_CONFIG['transportasi']['template']['sample']

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
