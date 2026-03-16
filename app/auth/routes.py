"""
Routes autentikasi — tanpa database.
"""
from flask import render_template, redirect, url_for, flash, request, current_app
from flask_login import login_user, logout_user, login_required, current_user
from werkzeug.security import check_password_hash
from app.auth import auth
from app.models import User


@auth.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('main.dia_brs'))

    if request.method == 'POST':
        username = request.form.get('username', '').strip().lower()
        password = request.form.get('password', '').strip()

        from app.services.sheets import get_users
        users_sheet = get_users()
        user_data = users_sheet.get(username)

        is_valid = False
        if user_data:
            # User ditemukan di GSheets -> cek plaintext password
            if user_data.get('password') == password:
                is_valid = True
        else:
            # Fallback ke lokal config
            user_data = current_app.config['USERS'].get(username)
            if user_data and check_password_hash(user_data['password_hash'], password):
                is_valid = True

        if is_valid:
            akses = {k: v for k, v in user_data.items() if k.startswith('akses_')}
            user = User(username=username, nama=user_data['nama'], akses=akses)
            login_user(user)
            return redirect(url_for('main.dia_brs'))

        flash('Username atau password salah.', 'danger')

    return render_template('auth/login.html')


@auth.route('/refresh_users')
def refresh_users():
    """Route untuk membersihkan cache data user dari Google Sheets."""
    from app.services.sheets import clear_users_cache
    clear_users_cache()
    flash('Data akun berhasil disinkronisasi ulang dengan Google Sheets.', 'success')
    return redirect(url_for('auth.login'))


@auth.route('/logout')
@login_required
def logout():
    logout_user()
    flash('Anda telah logout.', 'info')
    return redirect(url_for('auth.login'))
