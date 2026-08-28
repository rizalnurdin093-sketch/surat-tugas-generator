"""End-to-end test via Flask test_client — deterministik, tanpa server eksternal."""
import os
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, "/home/ubuntu/.local/lib/python3.12/site-packages")

from app import app
from docx import Document

# Gunakan folder temp untuk upload/generate
app.config["TESTING"] = True
tmp = tempfile.mkdtemp()
app.config["UPLOAD_FOLDER"] = os.path.join(tmp, "uploads")
app.config["GENERATED_FOLDER"] = os.path.join(tmp, "generated")
os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
os.makedirs(app.config["GENERATED_FOLDER"], exist_ok=True)

SRC = "uploads/260826 Daftar Peserta Pengelolaan SDM IHT.xlsx"
DEST = os.path.join(app.config["UPLOAD_FOLDER"], os.path.basename(SRC))
import shutil
shutil.copy(SRC, DEST)

with app.test_client() as c:
    # 1. Login
    r = c.post("/login", data={"username": "admin", "password": "admin123"})
    print("1. Login:", r.status_code, "(expect 302)")

    # 2. Preview (upload)
    with open(DEST, "rb") as f:
        r = c.post("/preview", data={
            "judul": "Pengelolaan SDM IHT",
            "tempat": "Ruang Meeting HC",
            "tanggal": "Senin-Kamis / 26-29 Agustus 2026",
            "tanggal_surat": "Agustus 2026",
            "file": (f, os.path.basename(SRC)),
        }, content_type="multipart/form-data")
    print("2. Preview:", r.status_code, "(expect 200)")
    html = r.get_data(as_text=True)
    print("   -> tombol generate ada?:", 'url_for' not in html and 'Generate Surat' in html)

    # 3. Session setelah preview
    with c.session_transaction() as s:
        print("3. Session uploaded_filename:", s.get("uploaded_filename"))
        print("   Session form:", s.get("form"))

    # 4. Generate
    r = c.post("/generate")
    print("4. Generate:", r.status_code)
    ct = r.headers.get("Content-Type", "")
    print("   Content-Type:", ct)
    if r.status_code == 200 and "spreadsheet" in ct or "octet" in ct or r.data[:2] == b"PK":
        # cek isi docx
        path = os.path.join(app.config["GENERATED_FOLDER"], "test.docx")
        with open(path, "wb") as f:
            f.write(r.data)
        doc = Document(path)
        print("   👍 DOCX VALID. Rows tabel peserta:", len(doc.tables[2].rows))
    else:
        with c.session_transaction() as s:
            print("   Flash:", s.get("_flashes"))
print("\nDONE")
