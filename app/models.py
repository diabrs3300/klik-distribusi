"""
User model sederhana — tanpa database.
"""
from flask_login import UserMixin
from app import login_manager


class User(UserMixin):
    def __init__(self, username, nama):
        self.id = username        # Flask-Login butuh atribut id
        self.username = username
        self.nama = nama


@login_manager.user_loader
def load_user(user_id):
    """Muat user dari config USERS berdasarkan username (id)."""
    from flask import current_app
    user_data = current_app.config['USERS'].get(user_id)
    if user_data:
        return User(username=user_id, nama=user_data['nama'])
    return None

