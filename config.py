"""
config.py — Konfigurasi aplikasi Surat Tugas
Semua setting yang bisa diubah (secret key, path folder, dll) di sini,
supaya tidak tersebar di banyak file dan gampang diubah.
"""

import os

BASE_DIR = os.path.abspath(os.path.dirname(__file__))


class Config:
    # Secret key untuk session flask-login.
    # GANTI ini dengan string acak di production (jangan dibiarkan default).
    SECRET_KEY = os.environ.get("SURAT_SECRET_KEY", "ganti-ini-dengan-secret-acak")

    # Database SQLite — satu file, gampang dibackup.
    # Kalau mau Postgres nanti, tinggal ganti driver + string ini.
    SQLALCHEMY_DATABASE_URI = f"sqlite:///{os.path.join(BASE_DIR, 'surat_tugas.db')}"

    # Folder-folder project
    UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
    GENERATED_FOLDER = os.path.join(BASE_DIR, "generated")
    TEMPLATE_SURAT_FOLDER = os.path.join(BASE_DIR, "templates_surat")

    # Ukuran maks file upload (15 MB)
    MAX_CONTENT_LENGTH = 15 * 1024 * 1024

    # Format file Excel yang diizinkan (hanya .xlsx — openpyxl tidak support .xls)
    ALLOWED_EXTENSIONS = {"xlsx"}
