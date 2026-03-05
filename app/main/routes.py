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
    'Tahun', 'Bulan', 'Kode Kota', 'Nama Kota', 'Kode Komoditas',
    'Nama Komoditas', 'Flag', 'NK', 'IHK',
    'Inflasi MtM', 'Inflasi YtD', 'Inflasi YoY',
    'Andil MtM', 'Andil YtD', 'Andil YoY',
]


@main.route('/upload-ihk', methods=['GET', 'POST'])
@login_required
def upload_ihk():
    if request.method == 'GET':
        return render_template('upload_ihk.html', title='Upload Excel IHK/Inflasi')

    # ── Ambil input form ──────────────────────────────────────────────────────
    bulan  = request.form.get('bulan', '').strip()
    tahun  = request.form.get('tahun', '').strip()
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
        from app.services.sheets import process_upload
        stats = process_upload(df, col_name)
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

