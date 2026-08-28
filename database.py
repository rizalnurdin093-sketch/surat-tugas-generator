"""
database.py — Semua interaksi database di satu file.
Supaya query SQL tidak tersebar di route app.py, jadi gampang dilacak
dan bisa dites terpisah.
"""

import os
import sqlite3
from werkzeug.security import generate_password_hash, check_password_hash

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "surat_tugas.db")


def get_db():
    """Buka koneksi database (baru tiap pemanggilan)."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Buat tabel-tabel kalau belum ada, dan seed user admin default."""
    conn = get_db()
    cur = conn.cursor()

    # Tabel user untuk login sederhana
    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # Tabel riwayat surat yang dibuat
    cur.execute("""
        CREATE TABLE IF NOT EXISTS surat_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            judul TEXT NOT NULL,
            tempat TEXT NOT NULL,
            tanggal TEXT NOT NULL,
            jumlah_peserta INTEGER DEFAULT 0,
            file_name TEXT NOT NULL,
            file_peserta TEXT,
            nomor_surat TEXT DEFAULT '',
            tanggal_surat TEXT DEFAULT '',
            dibuat_oleh TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # Migrasi ringan: tambah kolom jika DB dibuat sebelum ada kolom ini.
    # SQLite tidak support IF NOT EXISTS untuk ADD COLUMN, jadi kita cek pragma.
    cols = {row[1] for row in cur.execute("PRAGMA table_info(surat_history)").fetchall()}
    if "file_peserta" not in cols:
        cur.execute("ALTER TABLE surat_history ADD COLUMN file_peserta TEXT")
        print("Migrasi: kolom file_peserta ditambah ke surat_history.")
    if "nomor_surat" not in cols:
        cur.execute("ALTER TABLE surat_history ADD COLUMN nomor_surat TEXT DEFAULT ''")
        print("Migrasi: kolom nomor_surat ditambah ke surat_history.")
    if "tanggal_surat" not in cols:
        cur.execute("ALTER TABLE surat_history ADD COLUMN tanggal_surat TEXT DEFAULT ''")
        print("Migrasi: kolom tanggal_surat ditambah ke surat_history.")

    conn.commit()

    # Seed user admin default kalau belum ada
    if get_user("admin") is None:
        # Password default: ganti setelah login pertama
        _create_user("admin", "admin123")
        print("Admin user 'admin' dibuat (password default: admin123).")

    conn.close()


def _create_user(username, password):
    """Buat user baru (helper internal)."""
    conn = get_db()
    try:
        conn.execute(
            "INSERT INTO users (username, password_hash) VALUES (?, ?)",
            (username, generate_password_hash(password)),
        )
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False
    finally:
        conn.close()


def get_user(username):
    """Ambil user berdasarkan username."""
    conn = get_db()
    row = conn.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
    conn.close()
    return row


def verify_password(user, password):
    """Cek password cocok atau tidak."""
    return check_password_hash(user["password_hash"], password)


def add_surat_history(judul, tempat, tanggal, jumlah_peserta, file_name, dibuat_oleh, file_peserta=None, nomor_surat="", tanggal_surat=""):
    """Simpan riwayat surat yang berhasil dibuat."""
    conn = get_db()
    conn.execute(
        """
        INSERT INTO surat_history
            (judul, tempat, tanggal, jumlah_peserta, file_name, file_peserta, nomor_surat, tanggal_surat, dibuat_oleh)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (judul, tempat, tanggal, jumlah_peserta, file_name, file_peserta, nomor_surat, tanggal_surat, dibuat_oleh),
    )
    conn.commit()
    conn.close()


def get_all_surat_history():
    """Ambil semua riwayat surat (terbaru dulu)."""
    conn = get_db()
    rows = conn.execute("SELECT * FROM surat_history ORDER BY created_at DESC").fetchall()
    conn.close()
    return rows


def get_surat_by_id(surat_id):
    """Ambil satu record surat berdasarkan id."""
    conn = get_db()
    row = conn.execute("SELECT * FROM surat_history WHERE id = ?", (surat_id,)).fetchone()
    conn.close()
    return row


def update_nomor_surat(surat_id, nomor_surat):
    """Update nomor surat (string input = angka saja, misal '2630')."""
    conn = get_db()
    conn.execute(
        "UPDATE surat_history SET nomor_surat = ? WHERE id = ?",
        (nomor_surat.strip(), surat_id),
    )
    conn.commit()
    conn.close()


def search_surat_history(q):
    """
    Cari surat berdasarkan judul / nomor_surat / tempat (LIKE match, case-insensitive).
    Kosong -> semua. Tanda % ditolak untuk keamanan LIKE.
    """
    conn = get_db()
    if not q:
        rows = conn.execute("SELECT * FROM surat_history ORDER BY created_at DESC").fetchall()
    else:
        pattern = f"%{q.strip()}%"
        rows = conn.execute(
            """
            SELECT * FROM surat_history
            WHERE judul LIKE ? OR nomor_surat LIKE ? OR tempat LIKE ?
            ORDER BY created_at DESC
            """,
            (pattern, pattern, pattern),
        ).fetchall()
    conn.close()
    return rows
