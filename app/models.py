"""
User model sederhana — tanpa database.
"""
from flask_login import UserMixin
from app import login_manager


class User(UserMixin):
    def __init__(self, username, nama, akses=None):
        self.id = username        # Flask-Login butuh atribut id
        self.username = username
        self.nama = nama
        self.akses = akses or {}


@login_manager.user_loader
def load_user(user_id):
    """Muat user dari Google Sheets (via get_users), fallback ke config USERS."""
    from flask import current_app
    from app.services.sheets import get_users
    
    # 1. Coba dari Google Sheets
    users_sheet = get_users()
    user_data = users_sheet.get(user_id)
    
    # 2. Fallback ke config lokal jika tidak ketemu (atau jika fetch gagal)
    if not user_data:
        user_data = current_app.config['USERS'].get(user_id)
        
    if user_data:
        akses = {k: v for k, v in user_data.items() if k.startswith('akses_')}
        return User(username=user_id, nama=user_data['nama'], akses=akses)
    return None

