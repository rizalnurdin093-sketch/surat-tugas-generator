# SUGEN — Surat Tugas Generator

Aplikasi web internal untuk **membuat Surat Tugas (assignment letter)** secara otomatis di Divisi Human Capital **PT TIMAH (Persero) Tbk**.

Input detail surat + upload daftar peserta Excel → aplikasi menghasilkan **1 file `.docx`** yang siap pakai, lengkap dengan tabel lampiran peserta yang otomatis diurutkan A–Z.

---

## Fitur

- **Login** — proteksi sederhana username & password untuk jaringan kantor
- **Buat Surat** — isi judul pelatihan, tempat/lokasi, tanggal pelatihan, tanggal surat
- **Upload Excel (.xlsx)** — ekstrak otomatis kolom NIK / Nama / Divisi
- **Generate `.docx`** — 1 file surat + tabel lampiran peserta (auto A–Z)
- **Master Data** — arsip semua surat yang pernah dibuat:
  - Set nomor surat setelah mendapat penomoran dari tim admin
  - Cari surat berdasarkan judul / nomor / tempat
  - Export Excel (laporan) dengan kolom: ID, Judul, Nomor, Tempat, Tanggal Mulai, Tanggal Selesai
- **UI Corporate Clean** — navbar ikon, dropdown profil, design system konsisten (`DESIGN.md`)

---

## Tech Stack

| Lapisan | Teknologi |
|---------|-----------|
| Backend | Python + Flask |
| Database | SQLite |
| Dokumen | `python-docx` (generate .docx) |
| Spreadsheet | `openpyxl` (baca/export Excel) |
| Frontend | HTML + Jinja2 + CSS (vanilla, tanpa framework) |

---

## Struktur Folder

```
surat-tugas/
├── app.py               # Route Flask utama
├── config.py            # Konfigurasi (secret key, folder, dll)
├── database.py          # Koneksi & helper SQLite
├── parser_excel.py      # Baca & urutkan data peserta dari Excel
├── generator_surat.py   # Logika isi field + buat tabel lampiran .docx
├── run_prod.py          # Entry point produksi (DispatcherMiddleware)
├── requirements.txt
├── templates/           # HTML (base, index, login, master)
├── static/css/          # Design system
├── templates_surat/     # Template surat .docx (sumber)
└── DESIGN.md            # Panduan design (color, font, spacing)
```

---

## Instalasi & Menjalankan

```bash
# 1. Buat virtual environment
python3 -m venv venv
source venv/bin/activate

# 2. Install dependensi
pip install -r requirements.txt

# 3. Atur secret key (wajib di production)
export SURAT_SECRET_KEY="ganti-dengan-string-acak"

# 4. Jalankan server produksi
python run_prod.py
```

Server berjalan di `http://localhost:5001`.

> **Deploy di belakang Nginx (subpath `/surat/`):** aplikasi memakai `DispatcherMiddleware` dengan `SCRIPT_NAME=/surat`, sehingga bisa di-proxy dari `http://domain/surat/` tanpa konfigurasi tambahan.

---

## Format File Excel Peserta

Hanya mendukung **`.xlsx`** (openpyxl tidak support `.xls` lama).

Kolom yang dibaca (dicocokkan otomatis):
- **NIK** ← kolom `Employee No.`
- **Nama** ← kolom `Employee Name`
- **Divisi** ← kolom `Division`

Baris difilter dari sheet `Peserta Fix` yang kolom `Peserta`-nya bernilai `Ya`, lalu diurutkan A–Z.

---

## Kredensial Login (default)

| Username | Password |
|----------|----------|
| `admin`  | `admin123` |

> UBAH password default ini di production.

---

## Deployment Saat Ini

Aplikasi ter-deploy di VPS dengan **Nginx** → reverse proxy ke Flask (`localhost:5001`) melalui path `/surat/`.

---

## Nota Desain

Panduan lengkap warna, tipografi, spacing, dan pola komponen ada di **[`DESIGN.md`](DESIGN.md)** — dijadikan single source of truth agar seluruh halaman konsisten.

---

*Dibuat untuk Divisi Human Capital PT TIMAH (Persero) Tbk — © 2026*
