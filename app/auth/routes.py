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
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()

        user_data = current_app.config['USERS'].get(username)
        if user_data and check_password_hash(user_data['password_hash'], password):
            user = User(username=username, nama=user_data['nama'])
            login_user(user)
            return redirect(url_for('main.dia_brs'))

        flash('Username atau password salah.', 'danger')

    return render_template('auth/login.html', title='Login')


@auth.route('/logout')
@login_required
def logout():
    logout_user()
    flash('Anda telah logout.', 'info')
    return redirect(url_for('auth.login'))
