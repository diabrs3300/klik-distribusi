# DIA BRS

**Digital Integration & Automation for Berita Resmi Statistik**  
Tim GADIS (Harga dan Distribusi) — BPS Provinsi Jawa Tengah

---

## Tentang Aplikasi

DIA BRS adalah portal internal berbasis Flask yang melayani dua fungsi utama:

| Halaman | URL | Akses |
|---------|-----|-------|
| **Klik Distribusi** | `/` | Publik |
| **DIA BRS Dashboard** | `/dia-brs` | Login |
| **BRS IHK-Inflasi** | `/brs/ihk` | Login |
| **BRS NTP dan Gabah** | `/brs/ntp` | Login |
| **BRS Transportasi** | `/brs/transportasi` | Login |
| **BRS Ekspor Impor** | `/brs/ekspor` | Login |
| **BRS Pariwisata** | `/brs/pariwisata` | Login |
| **Upload Excel IHK** | `/upload-ihk` | Login |

---

## Struktur Project

```
dia-brs/
├── app.py                    # Entry point (python app.py)
├── config.py                 # Konfigurasi Flask + daftar akun
├── requirements.txt          # Dependencies produksi (tanpa pandas/openpyxl)
├── requirements-dev.txt      # Dependencies lokal (+ pandas, openpyxl)
├── vercel.json               # Konfigurasi deploy Vercel
├── credentials.json          # Service account Google (JANGAN di-commit!)
├── api/
│   └── index.py              # Entry point Vercel serverless
└── app/
    ├── constants.py          # ID spreadsheet & konfigurasi cache
    ├── models.py             # User model (tanpa database)
    ├── forms.py              # LoginForm (Flask-WTF)
    ├── errors.py             # Error handler 404/500
    ├── utils.py              # Helper functions
    ├── auth/                 # Blueprint autentikasi
    │   └── routes.py
    ├── main/                 # Blueprint halaman utama & BRS
    │   └── routes.py
    ├── services/
    │   └── sheets.py         # Koneksi & operasi Google Sheets
    ├── static/
    │   ├── css/style.css
    │   ├── js/main.js
    │   └── img/
    │       ├── favicon.png   # Favicon (kotak, 88x88px)
    │       └── logo-bps.webp
    └── templates/
        ├── base.html         # Layout utama (navbar, flash, footer)
        ├── index.html        # Klik Distribusi (standalone, tanpa base)
        ├── dashboard.html    # Dashboard DIA BRS
        ├── upload_ihk.html   # Form upload Excel IHK
        ├── auth/
        │   └── login.html
        ├── brs/              # Halaman per section BRS
        │   ├── ihk.html
        │   ├── ntp.html
        │   ├── transportasi.html
        │   ├── ekspor.html
        │   └── pariwisata.html
        └── errors/
            ├── 404.html
            └── 500.html
```

---

## Setup Lokal

### 1. Clone & Install

```bash
git clone <repo-url>
cd dia-brs
python -m venv venv
venv\Scripts\activate          # Windows
pip install -r requirements-dev.txt
```

### 2. Konfigurasi

Salin `.env.example` menjadi `.env`:

```bash
cp .env.example .env
```

Isi nilai berikut di `.env`:

```env
SECRET_KEY=isi-dengan-string-acak-yang-kuat
GOOGLE_CREDENTIALS_JSON={"type":"service_account",...}   # 1 baris JSON
```

> **Catatan:** `GOOGLE_CREDENTIALS_JSON` harus dalam **satu baris** JSON penuh.  
> Alternatif: letakkan `credentials.json` di root project (akan digunakan sebagai fallback).

### 3. Konfigurasi Spreadsheet

Edit `app/constants.py`:

```python
DOCS_SPREADSHEET_ID = 'ID-spreadsheet-BRS'        # untuk BRS docs
KLIK_SPREADSHEET_ID = 'ID-spreadsheet-Klik'       # untuk Klik Distribusi
DOCS_CACHE_TTL      = 300                          # cache TTL dalam detik
```

Pastikan kedua spreadsheet sudah di-share ke service account email (lihat `credentials.json` → `client_email`).

### 4. Jalankan

```bash
python app.py
```

Aplikasi berjalan di `http://localhost:5000`.

---

## Format Google Sheets

### Spreadsheet Klik Distribusi (sheet: `Klik`)

| Kolom | Isi |
|-------|-----|
| A | Kategori |
| B | Nama Tautan |
| C | URL |
| D | Jenis (Google Sheets / Web App / dll.) |
| E | Keyword (untuk pencarian) |
| F | Keterangan |

### Spreadsheet BRS (sheet: `IHK`, `NTP`, `Transportasi`, `Ekspor`, `Pariwisata`)

| Kolom | Isi |
|-------|-----|
| A | Nama Dokumen |
| B | Deskripsi |
| C | Link / URL |
| D | Kategori (`Master` / `BRS` / `Folder`) |

---

## Manajemen Akun

Tidak ada database — akun dikelola di `config.py`:

```python
USERS = {
    'username1': {
        'password_hash': generate_password_hash('password-disini'),
        'nama': 'Nama Tampil',
    },
}
```

---

## Cache & Refresh

Data dari Google Sheets di-cache selama `DOCS_CACHE_TTL` detik (default 5 menit).  
Untuk memaksa refresh data tanpa menunggu TTL:

- **Klik Distribusi:** klik tombol **Refresh** di halaman utama (`/klik-refresh`)
- **BRS Docs:** klik tombol **Refresh** di halaman BRS mana pun (`/docs-refresh`, perlu login)

---

## Deploy ke Vercel

1. Push ke GitHub
2. Hubungkan repo ke Vercel
3. Tambahkan **Environment Variables** di Vercel Dashboard:

| Key | Value |
|-----|-------|
| `GOOGLE_CREDENTIALS_JSON` | Isi JSON satu baris dari `credentials.json` |
| `SECRET_KEY` | String acak yang kuat |

> **Penting:** `pandas` dan `openpyxl` **tidak tersedia** di Vercel (batas ukuran).  
> Fitur Upload Excel hanya bisa digunakan di environment lokal.

---

## Layanan Sinkronisasi Google Sheets (`app/services/sheets.py`)

Aplikasi ini menggunakan modul terpusat `sheets.py` untuk mengelola koneksi dan operasi CRUD (*Create, Read, Update, Delete*) data dari dan ke Google Sheets API via autentikasi *Service Account*. 

Proses sinkronisasi data dipetah menjadi beberapa fungsi utama:

- **`_get_spreadsheet(spreadsheet_id)`**
  Pengelola otentikasi sentral. Mengambil kredensial klien *Service Account* (dari *environtment variables* atau `credentials.json`) dan membuka akses langsung ke Spreadsheet target.
  
- **`process_ihk_upload(dataframe, col_name)`**
  Mengorkestrasi seluruh sinkronisasi pengunggahan data **IHK**. Membaca file DataFrame (dari Excel) dan melakukan pembaruan lintas sheet secara *batching* sesuai pemetaan parameter dari BPS (misal ke sheet: *MTM*, *YTD*, *YOY*, dsb).
  
- **`process_ekspor_upload(dataframe)`**
  Mengorkestrasi sinkronisasi data agregasi **Ekspor**. Mengelompokkan nominal agregat (*sum*) **FOB** berdasarkan parameter *KODE_HS*, *PODAL5*, dan *NEWCTRYCOD* lalu mendistribusikannya ke sheet target masing-masing (*Barang*, *Pelabuhan*, *Negara*). Fungsi ini otomatis menambahkan header kolom periode *YYMM* baru (*auto-expanding*) jika periode belum terdaftar.

- **`_upsert_ihk_sheet` & `_upsert_ekspor_sheet`**
  Fungsi utilitas spesifik yang secara harfiah menangani manipulasi baris dan sel data mentah ke *worksheet*. Bertugas memetakan baris referensi via referensi silang (*lookup dict*), melakukan pembaruan *cell* (*update value*), dan/atau menempelkan baris ekstraksi jika item / komoditas tersebut belum ada sebelumnya (*appending row*).

---

## Dependencies

| Package | Kegunaan |
|---------|---------|
| Flask | Web framework |
| Flask-Login | Autentikasi sesi |
| Flask-WTF | Form + CSRF protection |
| gspread | Google Sheets API client |
| google-auth | Service account credentials |
| python-dotenv | Load `.env` |
| pandas *(dev)* | Baca file Excel (.xlsx) |
| openpyxl *(dev)* | Generate template Excel |
