"""
Entry point untuk Vercel Serverless Function.
Vercel membutuhkan objek `app` (WSGI) yang diekspos di sini.
"""
import sys
import os

# Pastikan root project ada di sys.path agar import `app` dan `config` berfungsi
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import create_app

app = create_app('production')
