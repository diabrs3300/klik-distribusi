"""
Script debug — jalankan dari root project:
    python debug_sheets.py
Untuk test koneksi ke Google Sheets dan melihat raw error.
"""
import sys, os, logging
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

logging.basicConfig(level=logging.DEBUG, format='%(levelname)s %(name)s: %(message)s')

from app.services.sheets import get_klik_links, get_docs

print("\n=== TEST: Klik Distribusi ===")
try:
    result = get_klik_links()
    if result:
        for g in result:
            print(f"  [{g['kategori']}] {len(g['links'])} item")
    else:
        print("  KOSONG / GAGAL (lihat error di atas)")
except Exception as e:
    print(f"  EXCEPTION: {e}")

print("\n=== TEST: Docs IHK ===")
try:
    result = get_docs('IHK', fallback=[])
    if result:
        for g in result:
            print(f"  [{g['kategori']}] {len(g['items'])} dokumen")
    else:
        print("  KOSONG / GAGAL (lihat error di atas)")
except Exception as e:
    print(f"  EXCEPTION: {e}")
