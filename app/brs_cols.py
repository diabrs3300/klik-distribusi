"""
Konfigurasi terpusat untuk kolom wajib, lembar kerja (sheets), dan
data sampel template Excel untuk Modul Upload BRS.

Konvensi untuk Ekspor & Impor:
  required_excel_cols : { SemanticKey: nama_kolom_di_Excel }
  required_dbf_cols   : { SemanticKey: nama_kolom_di_DBF  }
  → Key-nya identik, sehingga sheets.py cukup panggil SemanticKey tanpa peduli format sumber.
"""

BRS_CONFIG = {
    # ─── IHK - Inflasi ────────────────────────────────────────────────────────
    'ihk': {
        'theme_color': '0D6EFD',  # Blue
        'keys': ['KeyKodeKota', 'KeyKodeKomoditas'],
        # SemanticKey → nama kolom default di Excel (bisa di-override user)
        'required_cols': {
            'KeyTahun':         'Tahun',
            'KeyBulan':         'Bulan',
            'KeyKodeKota':      'Kode Kota',
            'KeyKodeKomoditas': 'Kode Komoditas',
            'KeyNK':            'NK',
            'KeyIHK':           'IHK',
            'KeyInflasiMtM':    'Inflasi MtM',
            'KeyInflasiYtD':    'Inflasi YtD',
            'KeyInflasiYoY':    'Inflasi YoY',
            'KeyAndilMtM':      'Andil MtM',
            'KeyAndilYtD':      'Andil YtD',
            'KeyAndilYoY':      'Andil YoY',
        },
        # SemanticKey → nama kolom di Google Sheets (fixed, tidak berubah)
        'mapping_cols': {
            'KeyKodeKota':      'Kd.Kota',
            'KeyKodeKomoditas': 'Kode',
            'KeyNK':            'NK',
            'KeyIHK':           'IHK',
            'KeyInflasiMtM':    'MTM',
            'KeyInflasiYtD':    'YTD',
            'KeyInflasiYoY':    'YOY',
            'KeyAndilMtM':      'AMTM',
            'KeyAndilYtD':      'AYTD',
            'KeyAndilYoY':      'AYOY',
        },
        'sheets': ['IHK', 'NK', 'MTM', 'YTD', 'YOY', 'AMTM', 'AYTD', 'AYOY'],
        'template': {
            'sample': [
                2026, 2, 3300, '01',
                45285204773250.40, 117.08,
                2.02, 0.22, 2.89,
                0.58, 0.06, 0.85,
            ],
            'col_widths': [8, 7, 15, 11, 20, 11, 12, 12, 12, 12, 12, 12],
        },
    },

    # ─── Ekspor ───────────────────────────────────────────────────────────────
    'ekspor': {
        'theme_color': '0D9488',  # Teal
        'required_excel_cols': {
            'KeyBulan':     'BLN_PROSES',
            'KeyTahun':     'THN_PROSES',
            'KeyPelabuhan': 'PELABUHAN',
            'KeyKodeHS':    'KODE_HS',
            'KeyNegara':    'NEGARA',
            'KeyNilai':     'FOB',
        },
        'required_dbf_cols': {
            'KeyBulan':     'MTH',
            'KeyTahun':     'YEAR',
            'KeyPelabuhan': 'PELABUHAN',
            'KeyKodeHS':    'KODE_HS',
            'KeyNegara':    'NEWCTRYCOD',
            'KeyNilai':     'FOB',
        },
        'targets': {
            'Pelabuhan':    'KeyPelabuhan',
            'NegaraBarang': ['KeyNegara', 'KeyKodeHS'],
        },
        'sheets': ['Pelabuhan', 'NegaraBarang', 'RECAP'],
        'template': {
            'sample': ['1', '2024', '33494', '01061100', '516', '1436.400'],
            'col_widths': [8, 6, 10, 10, 14, 12],
        },
    },

    # ─── Impor ───────────────────────────────────────────────────────────────
    'impor': {
        'theme_color': '1D4ED8',  # Dark Blue
        'required_excel_cols': {
            'KeyBulan':   'BLN',
            'KeyTahun':   'THN_PROSES',
            'KeyKodeHS':  'KODE_HS',
            'KeyNegara':  'NEGARA',
            'KeyNilai':   'NILAI',
        },
        'required_dbf_cols': {
            'KeyBulan':   'MTH',
            'KeyTahun':   'YEAR',
            'KeyKodeHS':  'HS',
            'KeyNegara':  'K_NEGARA',
            'KeyNilai':   'NILAI',
        },
        'targets': {
            'NegaraBarang': ['KeyNegara', 'KeyKodeHS'],
        },
        'sheets': ['NegaraBarang', 'RECAP'],
        'template': {
            'sample': ['1', '2024', '32129029', '516', '1186.0'],
            'col_widths': [8, 12, 16, 14, 16, 14],
        },
    },

    # ─── NTP ──────────────────────────────────────────────────────────────────
    'ntp': {
        'theme_color': '198754',  # Green
        'required_cols': {
            'KeyTahun': 'Tahun', 
            'KeyBulan': 'Bulan',
            'KeyKdProv': 'Kd.Prov', 
            'KeyNamaProv': 'Nama Prov',
            'KeyNTP': 'NTP', 
            'KeyNTPP': 'NTPP', 
            'KeyNTPH': 'NTPH',
            'KeyNTPN': 'NTPN', 
            'KeyNTPF': 'NTPF', 
            'KeyNTPE': 'NTPE',
        },
        'sheets': ['NTP', 'NTPP', 'NTPH', 'NTPN', 'NTPF', 'NTPE'],
        'template': {
            'sample': ['2026', '8', '33', 'JAWA TENGAH', '105.12', '102.30', '107.50', '103.00', '108.20', '104.10'],
            'col_widths': [8, 7, 10, 15, 10, 10, 10, 10, 10, 10],
        },
    },

    # ─── Pariwisata ───────────────────────────────────────────────────────────
    'pariwisata': {
        'theme_color': '7C3AED',  # Purple
        'required_cols': {
            'KeyTahun': 'Tahun', 
            'KeyBulan': 'Bulan',
            'KeyNamaHotel': 'Nama Hotel', 
            'KeyKelasHotel': 'Kelas Hotel',
            'KeyTPK': 'TPK', 
            'KeyRATS': 'RATS',
            'KeyTamuAsing': 'Tamu Asing', 
            'KeyTamuDomestik': 'Tamu Domestik',
        },
        'sheets': ['TPK', 'RATS', 'Tamu'],
        'template': {
            'sample': ['2026', '8', 'Grand Artos', 'Bintang 4', '62.50', '2.10', '120', '850'],
            'col_widths': [8, 7, 18, 14, 10, 10, 14, 14],
        },
    },

    # ─── Transportasi ─────────────────────────────────────────────────────────
    'transportasi': {
        'theme_color': 'D97706',  # Orange
        'required_cols': {
            'KeyTahun': 'Tahun', 
            'KeyBulan': 'Bulan',
            'KeyModa': 'Moda', 
            'KeyNamaPerusahaan': 'Nama Perusahaan',
            'KeyPenumpangBerangkat': 'Penumpang Berangkat',
            'KeyPenumpangDatang': 'Penumpang Datang',
            'KeyBarangMuat': 'Barang Muat', 
            'KeyBarangBongkar': 'Barang Bongkar',
        },
        'sheets': ['Udara', 'Laut', 'KA', 'ASDP'],
        'template': {
            'sample': ['2026', '8', 'Udara', 'Garuda Indonesia', '12500', '11800', '850.50', '920.30'],
            'col_widths': [8, 7, 12, 18, 18, 18, 18, 18],
        },
    },
}
