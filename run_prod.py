"""run_prod.py — Jalankan app di subpath /surat/ untuk behind Nginx."""
from app import create_dispatched_app

if __name__ == "__main__":
    create_dispatched_app()
