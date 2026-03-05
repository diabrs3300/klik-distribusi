"""
constants.py — Konfigurasi registry dokumen BRS via Google Sheets.
"""

# ─── Registry Dokumen BRS via Google Sheets ───────────────────────────────────
DOCS_SPREADSHEET_ID = '1avsEtTX-mNJ5_vSFt2cMT7nR8dporQviNKXoS20EUkw'  # ← ID spreadsheet BRS

# Mapping: nama section → nama sheet di Google Sheets registry
DOCS_SHEET_MAP = {
    'IHK'         : 'IHK',
    'NTP'         : 'NTP',
    'Transportasi': 'Transportasi',
    'Ekspor'      : 'Ekspor',
    'Pariwisata'  : 'Pariwisata',
}

# ─── Klik Distribusi via Google Sheets ────────────────────────────────────────
KLIK_SPREADSHEET_ID = '1GqZ9a95FuNLsqEphO2aBzc0ZYlVIzhpk3_ltG-wSaII'  # ← ID spreadsheet Klik Distribusi

# ─── TTL Cache (berlaku untuk keduanya) ───────────────────────────────────────
DOCS_CACHE_TTL = 300  # detik (default: 5 menit)
