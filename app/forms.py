"""
forms.py — Form definitions menggunakan Flask-WTF.
Saat ini hanya LoginForm yang aktif digunakan (tidak ada registrasi publik).
"""
from flask_wtf import FlaskForm
from wtforms import StringField, PasswordField, BooleanField, SubmitField
from wtforms.validators import DataRequired


class LoginForm(FlaskForm):
    """Form login sederhana menggunakan username + password."""
    username = StringField('Username', validators=[DataRequired()])
    password = PasswordField('Password', validators=[DataRequired()])
    remember_me = BooleanField('Ingat Saya')
    submit = SubmitField('Masuk')
