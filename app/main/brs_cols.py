"""
Konfigurasi terpusat untuk kolom wajib, lembar kerja (sheets), dan
data sampel template Excel untuk Modul Upload BRS.
"""

BRS_CONFIG = {
    'ihk': {
        'required_cols': [
            'Tahun', 'Bulan', 'Kode Kota', 'Kode Komoditas',
            'NK', 'IHK',
            'Inflasi MtM', 'Inflasi YtD', 'Inflasi YoY',
            'Andil MtM', 'Andil YtD', 'Andil YoY',
        ],
        'optional_cols': ['Nama Kota', 'Nama Komoditas', 'Flag'],
        'sheets': ['IHK', 'NK', 'MTM', 'YTD', 'YOY', 'AMTM', 'AYTD', 'AYOY'],
        'template': {
            'sample': [
                2026, 2, 3300, '01',
                45285204773250.40, 117.08,
                2.02, 0.22, 2.89,
                0.58, 0.06, 0.85
            ],
            'col_widths': [8, 7, 15, 11, 20, 11, 12, 12, 12, 12, 12, 12],
        }
    },
    'ekspor': {
        'required_cols': [
            'YEAR', 'MTH', 'KODE_HS', 'PODAL5',
            'NEWCTRYCOD', 'FOB'
        ],
        'dbf_cols': [
            'YEAR', 'MTH', 'KODE_HS', 'PODAL5',
            'NEWCTRYCOD', 'FOB'
        ],
        'excel_cols': [
            'BLN_PROSES', 'THN_PROSES', 'PELABUHAN', 
            'KODE_HS', 'NEGARA', 'FOB'
        ],
        'template': {
            'sample': ['1', '2024', '33494', '01061100', '516', '1436.400'],
            'col_widths': [8, 6, 10, 10, 14, 12],
        }
    },
    'impor': {
        'required_cols': [
            'HS', 'K_NEGARA'
        ],
        'dbf_cols': [
            'HS', 'K_NEGARA', 'N[Bulan][Tahun]'
        ],
        'excel_cols': [
            'BLN', 'THN_PROSES', 'KODE_HS', 'NEGARA', 'NILAI'
        ],
        'template': {
            'sample': ['1', '2024', '32129029', '516', '1186.0'],
            'col_widths': [8, 12, 16, 14, 16, 14],
        }
    },
    'ntp': {
        'required_cols': [
            'Tahun', 'Bulan', 'Kode Provinsi', 'Nama Provinsi',
            'NTP', 'NTPP', 'NTPH', 'NTPN', 'NTPF', 'NTPE',
        ],
        'sheets': ['NTP', 'NTPP', 'NTPH', 'NTPN', 'NTPF', 'NTPE'],
        'template': {
            'sample': ['2026', '8', '33', 'JAWA TENGAH', '105.12', '102.30', '107.50', '103.00', '108.20', '104.10'],
            'col_widths': [14] * 10,
        }
    },
    'pariwisata': {
        'required_cols': [
            'Tahun', 'Bulan', 'Nama Hotel', 'Kelas Hotel',
            'TPK', 'RATS', 'Tamu Asing', 'Tamu Domestik',
        ],
        'sheets': ['TPK', 'RATS', 'Tamu'],
        'template': {
            'sample': ['2026', '8', 'Grand Artos', 'Bintang 4', '62.50', '2.10', '120', '850'],
            'col_widths': [16] * 8,
        }
    },
    'transportasi': {
        'required_cols': [
            'Tahun', 'Bulan', 'Moda', 'Nama Perusahaan',
            'Penumpang Berangkat', 'Penumpang Datang', 'Barang Muat', 'Barang Bongkar',
        ],
        'sheets': ['Udara', 'Laut', 'KA', 'ASDP'],
        'template': {
            'sample': ['2026', '8', 'Udara', 'Garuda Indonesia', '12500', '11800', '850.50', '920.30'],
            'col_widths': [18] * 8,
        }
    }
}
