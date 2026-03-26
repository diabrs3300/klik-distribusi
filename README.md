# DIA BRS

**Digital Integration & Automation for Berita Resmi Statistik**  
Tim GADIS (Harga dan Distribusi) — BPS Provinsi Jawa Tengah

---

### 👥 Tim Pengembang

| Role | Developer | Kontribusi Utama |
| :--- | :--- | :--- |
| **Project Lead & Web Developer** | **D - Dutatama Rosewika Taufiq Hadihardaya** | Bertanggung jawab atas perencanaan, pengembangan, dan integrasi sistem end-to-end. Mengembangkan web app serta menghubungkan backend dengan layanan Google Sheets melalui Google Cloud API untuk mendukung alur data dan automasi sistem. |
| **Document Integration Engineer** | **I - Isnatul Mu'anissah** | Mengintegrasikan hasil pengolahan data ke dalam berbagai format dokumen, infografis, dan media presentasi (Docs, Slides, Drawing) secara otomatis, serta memastikan pembaruan data berjalan sinkron dan real-time. |
| **Data Processing Engineer (Google Sheets Backend)** | **A - Azmi Zulfani Putri** | Menangani logika pemrosesan data di Google Sheets, termasuk pengolahan, transformasi, tabulasi, dan visualisasi data untuk menghasilkan output yang terstruktur dan siap digunakan. |

---

## 🌟 Tentang Aplikasi

DIA BRS adalah platform internal berbasis Flask yang dirancang untuk mengintegrasikan manajemen dokumen dan otomatisasi pengolahan data Berita Resmi Statistik (BRS) melalui ekosistem Google Sheets.

Aplikasi ini memiliki dua komponen utama:
1.  **Klik Distribusi** (`/`): Portal publik untuk akses cepat tautan layanan internal dan eksternal.
2.  **Dashboard DIA BRS** (`/dia-brs`): Area terbatas untuk mengelola dokumen BRS (IHK, NTP, Ekspor-Impor, dll.) dan melakukan pengunggahan data mentah ke basis data Google Sheets.

---

## 🏗️ Struktur Proyek

```text
dia-brs/
├── app.py                  # Entry point aplikasi Flask
├── config.py               # Konfigurasi Flask & user management (fallback)
├── constants.py            # Mapping ID spreadsheet & konfigurasi cache
├── requirements.txt        # Daftar dependensi utama
├── vercel.json             # Konfigurasi deployment serverless Vercel
├── api/                    # Handler untuk Vercel Serverless Functions
└── app/
    ├── brs_cols.py         # Konfigurasi semantik kolom & tema per section BRS
    ├── constants.py        # ID spreadsheet & konfigurasi cache (link path)
    ├── models.py           # User model & logic autentikasi
    ├── services/
    │   └── sheets.py       # Engine utama integrasi Google Sheets API
    ├── static/             # Assets (CSS/JS/Images)
    └── templates/
        ├── base.html       # Layout utama dengan sidebar & navbar
        ├── index.html      # Landing page Klik Distribusi
        ├── dashboard.html  # Menu utama DIA BRS
        ├── brs/            # List halaman BRS (Halaman detail & Form upload)
        │   ├── ihk.html | upload_ihk.html
        │   ├── ntp.html | upload_ntp.html
        │   ├── transportasi.html | upload_transportasi.html
        │   ├── ekspor_impor.html | upload_ekspor_impor.html
        │   └── pariwisata.html | upload_pariwisata.html
        └── errors/         # Halaman penanganan error (404, 500)
```

---

## 🚀 Fitur Unggulan

### 1. Document Registry Sychronization
Manajemen tautan dokumen (Naskah BRS, Master File, Database) dilakukan sepenuhnya melalui Google Sheets. Perubahan di spreadsheet akan langsung tercermin di web setelah cache di-refresh. Mendukung layout **4-kolom** otomatis dan kategori dokumen: `Master`, `BRS`, `Database`, dan `Folder`.

### 2. Intelligent Data Upload (Excel/DBF)
Modul upload cerdas untuk **IHK**, **Ekspor**, dan **Impor** yang mendukung:
- **Semantic Mapping**: Mencocokkan kolom file sumber (.xlsx/.dbf) dengan kolom target di Google Sheets menggunakan konfigurasi di `brs_cols.py`.
- **Auto-Expanding Rows/Cols**: Otomatis menambah baris komoditas baru atau kolom periode bulan/tahun baru.
- **Batch Updates**: Mengirimkan ribuan data dalam satu request untuk performa optimal.

### 3. Granular Access Control
Akses ke setiap section BRS (IHK, NTP, Ekspor-Impor, dsb.) dikelola per user. Pengaturan akun dapat dilakukan melalui Google Sheets (sheet `akun`) atau via `config.py`.

---

## 🛠️ Setup & Instalasi Lokal

1.  **Persiapan environment**:
    ```bash
    python -m venv venv
    venv\Scripts\activate
    pip install -r requirements.txt
    ```

2.  **Konfigurasi Environment (`.env`)**:
    Buat file `.env` di root folder:
    ```env
    SECRET_KEY=kuncirahasiaanda
    GOOGLE_CREDENTIALS_JSON={"type":"service_account", ...} # JSON 1 baris
    ```
    *Atau letakkan file `credentials.json` di root folder.*

3.  **Spreadsheet IDs**:
    Pastikan Spreadsheet ID di `app/constants.py` sudah sesuai dan file tersebut telah di-share ke email **Service Account**.

4.  **Menjalankan Server**:
    ```bash
    python app.py
    ```

---

## 📊 Skema Google Sheets

### 1. Registry Dokumen BRS
Setiap section (IHK, NTP, dsb.) memiliki sheet dengan format 4 kolom:
- **Nama Dokumen**: Judul tautan yang tampil.
- **Kategori**: `Master`, `BRS`, `Database`, atau `Folder`.
- **Deskripsi**: Penjelasan singkat dokumen.
- **Link**: URL tujuan (Google Drive/Docs/Sheets).

### 2. Klik Distribusi
Sheet `Klik` pada spreadsheet Klik Distribusi:
- **Kategori**: Pengelompokan tautan.
- **Nama**: Label tautan.
- **Link**: URL tujuan.
- **Jenis**: Desktop, Web App, dsb (menentukan icon).
- **Tahun & Keterangan**: Metadata tambahan.

---

## ☁️ Deployment (Vercel)

Aplikasi ini siap dideploy ke Vercel. 
> **Catatan Penting**: Library `pandas`, `openpyxl`, dan `dbfread` memiliki ukuran yang besar. Jika dideploy ke environment serverless dengan batasan ukuran (seperti Vercel tingkat gratis), perhatikan batas ukuran package yang diizinkan.

---

## 📄 Lisensi
Tim GADIS - BPS Provinsi Jawa Tengah. Internal Use Only.
