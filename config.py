"""
Konfigurasi aplikasi Flask.
Tidak menggunakan database — user disimpan di dictionary.
"""
import os
from werkzeug.security import generate_password_hash
from dotenv import load_dotenv

load_dotenv()


class Config:
    """Konfigurasi dasar."""
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'kunci-rahasia-dia-brs-2026'
    DEBUG = False
    TESTING = False

    # Daftar akun yang diizinkan login (tanpa database)
    USERS = {
        'diabrs3300': {
            'password_hash': generate_password_hash('youngcc2026'),
            'nama': 'DIA BRS 3300',
        },
        'duta0000': {
            'password_hash': generate_password_hash('duta0000'),
            'nama': 'Duta',
        },
    }


class DevelopmentConfig(Config):
    DEBUG = True


class TestingConfig(Config):
    TESTING = True


class ProductionConfig(Config):
    DEBUG = False


config = {
    'development': DevelopmentConfig,
    'testing': TestingConfig,
    'production': ProductionConfig,
    'default': DevelopmentConfig,
}

