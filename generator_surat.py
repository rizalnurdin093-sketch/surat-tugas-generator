"""
generator_surat.py — Generate surat tugas .docx dari template.

Logic:
- Buka template_surat.docx
- Isi field: Hari/Tanggal (P10), Judul (P11), Tempat (P12), judul daftar peserta (P35)
- Kosongkan (sesuai keputusan user): Nomor (P1), TTD (Tabel 0), Nomor & Tanggal Lampiran (Tabel 1)
- Hapus baris kosong di Tabel 2, ganti dengan baris peserta sesuai jumlah (auto-skalasi)
- Output: file .docx siap download

Dipisah dari app.py supaya gampang dites & dilacak kalau ada bug.
"""

import os
import shutil
import re
from copy import deepcopy

from docx import Document


TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "templates_surat", "template_surat.docx")

# Index yang sudah diverifikasi dari inspeksi template.
P_HARI_TANGGAL = 10
P_JUDUL = 11
P_TEMPAT = 12
P_JUDUL_LAMPIRAN = 35

TABLE_TTD = 0          # Baris TTD dibiarkan seperti template (PT TIMAH, dll terisi)
TABLE_LAMPIRAN = 1     # Nomor + Tanggal lampiran dikosongkan (nomor manual)
TABLE_PESERTA = 2      # Daftar peserta — auto-skalasi


def _clean(text):
    """Bersihkan karakter non-breaking space dan rapikan whitespace."""
    if text is None:
        return ""
    return " ".join(str(text).replace("\xa0", " ").split())


def _replace_in_paragraph(paragraph, old_text, new_text):
    """
    Ganti old_text dengan new_text di dalam paragraph dengan MEMPERTAHANKAN format
    run pertama yang mengandung teks tsb (bukan menulis ulang ke run 0).

    Digunakan untuk kolom label+placeholder yang format bold-nya harus dijaga.
    """
    remaining = old_text
    for idx, run in enumerate(paragraph.runs):
        if not remaining:
            # teks target sudah habis, kosongkan sisa run
            run.text = ""
            continue
        if run.text and remaining.startswith(run.text):
            remaining = remaining[len(run.text):]
            continue
        # run ini adalah bagian dari teks yang akan diganti
        if new_text and run.text:
            run.text = new_text
            new_text = ""  # hanya isi di satu run
        else:
            run.text = ""


def _set_field_bold_placeholder(paragraph, search_placeholder, new_value):
    """
    Ganti teks placeholder (yang BOLD) di paragraph dengan new_value,
    mempertahankan format run tsb. Mencari run yang mengandung placeholder
    DAN bold=True (karena label di P10/P11/P12 non-bold, nilai placeholder bold).
    """
    # Prioritas: run yang contains placeholder AND bold
    for run in paragraph.runs:
        if run.bold and search_placeholder in run.text and run.text.strip():
            run.text = new_value
            return True
    # Fallback: run pertama yang contains placeholder
    for run in paragraph.runs:
        if search_placeholder in run.text and run.text.strip():
            run.text = new_value
            return True
    return False


def _set_cell_first_run(cell, new_value, clear_other_paragraphs=True):
    """
    Isi nilai pada run pertama paragraph pertama di cell, kosongkan sisa runs.
    Mempertahankan format (bold/size) run pertama.
    """
    if clear_other_paragraphs:
        for p in cell.paragraphs[1:]:
            p._element.getparent().remove(p._element)
    first_p = cell.paragraphs[0]
    if first_p.runs:
        first_p.runs[0].text = new_value
        for run in first_p.runs[1:]:
            run.text = ""
    else:
        first_p.add_run(new_value)


def _append_to_paragraph(paragraph, suffix):
    """Append text di akhir paragraph (ke run terakhir)."""
    if paragraph.runs:
        paragraph.runs[-1].text += suffix
    else:
        paragraph.add_run(suffix)


def _remove_table_rows(table, n_rows_to_remove):
    """
    Hapus baris dari bawah tabel sebanyak n_rows_to_remove.
    python-docx tidak punya remove_row built-in, jadi pakai XML manipulation.
    """
    tbl = table._tbl
    rows = list(tbl.findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr"))
    for row in rows[-n_rows_to_remove:]:
        tbl.remove(row)


def _add_row_like(table, template_row):
    """
    Tambah baris baru ke tabel dengan menyalin XML baris template (format sama),
    lalu hapus isi teksnya.

    python-docx Table.add_row() tidak bisa clone format, jadi kita salin XML
    <w:tr> secara manual dan kosongkan cell-nya.
    """
    import copy
    from docx.oxml.ns import qn

    new_tr = copy.deepcopy(template_row._tr)
    table._tbl.append(new_tr)

    # Kosongkan teks semua cell di baris baru (biar nanti diisi data)
    from docx.table import _Row
    new_row = _Row(new_tr, table)
    for cell in new_row.cells:
        _clear_cell_text(cell)
    return new_row


def _clear_cell_text(cell):
    """Kosongkan teks di cell, sisakan satu paragraph kosong."""
    # Hapus semua paragraph kecuali yang pertama, kosongkan yang pertama
    for p in cell.paragraphs[1:]:
        p._element.getparent().remove(p._element)
    first = cell.paragraphs[0]
    for run in list(first.runs):
        run.text = ""


FONT_TABEL = "Arial"
FONT_TABEL_SIZE = 10  # pt


def _apply_cell_font(cell):
    """Paksa font Arial 10 di SEMUA run paragraph cell (tabel peserta)."""
    from docx.shared import Pt
    for p in cell.paragraphs:
        for run in p.runs:
            run.font.name = FONT_TABEL
            run.font.size = Pt(FONT_TABEL_SIZE)


def _fill_peserta_row(row, no, nama, nik, divisi):
    """Isi satu baris tabel peserta dengan NO, NAMA, NIK, DIVISION (font Arial 10)."""
    cells = row.cells
    if len(cells) < 4:
        return
    values = [str(no), _clean(nama), _clean(nik), _clean(divisi)]
    for cell, val in zip(cells, values):
        # Hapus dulu semua paragraph di cell (kecuali paragraph pertama)
        for p in cell.paragraphs[1:]:
            p._element.getparent().remove(p._element)
        # Set teks di paragraph pertama
        first_p = cell.paragraphs[0]
        if first_p.runs:
            first_p.runs[0].text = val
            for run in first_p.runs[1:]:
                run.text = ""
        else:
            first_p.add_run(val)
        # Paksa font Arial 10
        _apply_cell_font(cell)


def generate_surat(
    judul: str,
    tempat: str,
    hari_tanggal: str,
    tanggal_surat: str,
    peserta: list,
    output_dir: str,
    nomor_angka: str = "",
) -> str:
    """
    Generate 1 file .docx surat tugas.
    Return path ke file hasil.

    Args:
        judul: judul pelatihan (cth: "Human Resources Management Essentials")
        tempat: lokasi pelatihan (cth: "Timah Learning Center - Pemali")
        hari_tanggal: tanggal pelatihan (cth: "Senin-Selasa / 31 Agustus - 01 September 2026")
        tanggal_surat: bulan+ tahun terbit surat (cth: "September 2026"). Hari tetap kosong manual.
        peserta: list of dict {'nama', 'nik', 'divisi'} — sudah urut A-Z
        nomor_angka: opsional — string angka (cth "2630") yang akan diletakkan di depan
                     "/Tbk/ST-4010/26-S8.7.1". Kosong -> template dibiarkan apa adanya.
        output_dir: folder tempat simpan file hasil

    Returns:
        (absolute_path, filename) file .docx yang berhasil dibuat
    """
    if not peserta:
        raise ValueError("Daftar peserta kosong, tidak bisa generate surat.")

    # Salin template ke output file (jangan modify file template langsung)
    safe_judul = re.sub(r"[^\w\s-]", "", judul).strip().replace(" ", "_")[:50]
    output_filename = f"surat_tugas_{safe_judul or 'tanpa_judul'}.docx"
    output_path = os.path.join(output_dir, output_filename)
    shutil.copyfile(TEMPLATE_PATH, output_path)

    doc = Document(output_path)

    # 1) Isi field Hari/Tanggal, Judul, Tempat (P10, P11, P12)
    #    Template baru punya placeholder BOLD yang diganti nilainya → format bold terjaga.
    p_hari = doc.paragraphs[P_HARI_TANGGAL]          # 'Hari/Tanggal\t: Hari/Tanggal\t' (nilai bold)
    _set_field_bold_placeholder(p_hari, "Hari/Tanggal", hari_tanggal)

    p_judul = doc.paragraphs[P_JUDUL]                # 'Judul\t\t: Judul' (nilai bold)
    _set_field_bold_placeholder(p_judul, "Judul", judul)

    p_tempat = doc.paragraphs[P_TEMPAT]              # 'Tempat\t\t: Tempat' (nilai bold)
    _set_field_bold_placeholder(p_tempat, "Tempat", tempat)

    # 2) Nomor surat (P1) & Tabel 1 (Lampiran):
    #    - Kalau nomor_angka kosong: biarkan template apa adanya ('/Tbk/...')
    #    - Kalau nomor_angka ada: jadi '2630/Tbk/ST-4010/26-S8.7.1'
    if nomor_angka:
        nomor_lengkap = f"{nomor_angka}/Tbk/ST-4010/26-S8.7.1"
        # P1
        p_nomor = doc.paragraphs[1]
        _replace_in_paragraph(p_nomor, "/Tbk/ST-4010/26-S8.7.1", nomor_lengkap)
        # Tabel 1 row1 cell2
        tabel_lampiran = doc.tables[TABLE_LAMPIRAN]
        sel_nomor = tabel_lampiran.rows[1].cells[2]
        _replace_in_paragraph(sel_nomor.paragraphs[0], "/Tbk/ST-4010/26-S8.7.1", nomor_lengkap)
    # else: biarkan template apa adanya (sudah '/Tbk/...')

    # 3) Blok TTD (Tabel 0): biarkan seperti template (PT TIMAH, jabatan, nama penandatangan).
    #    Yang diisi dari form: "Pada tanggal : <bulan> <tahun>" (hari tetap kosong manual).
    tabel_ttd = doc.tables[TABLE_TTD]
    if tanggal_surat:
        sel_tanggal = tabel_ttd.rows[0].cells[0]
        for p in sel_tanggal.paragraphs:
            if "Pada tanggal" in p.text:
                # p.text: "Pada tanggal  :       Agustus 2026" → ganti 'Agustus' dengan input,
                # sisakan spasi di depan untuk penulisan hari manual.
                _set_field_bold_placeholder(p, "Agustus", tanggal_surat.split()[-2] if len(tanggal_surat.split())>1 else tanggal_surat)

    # 4) Tabel 1 (Lampiran): Nomor DIBIARKAN seperti template (ada '/Tbk/...').
    #    Tanggal: isi dari tanggal_surat (bulan+tahun), diberi spasi depan utk nulis hari.
    tabel_lampiran = doc.tables[TABLE_LAMPIRAN]
    # row2 cell2 = nilai tanggal (bold, ukuran template dijaga)
    sel_lamp_tanggal = tabel_lampiran.rows[2].cells[2]
    _set_cell_first_run(sel_lamp_tanggal, f"{tanggal_surat}", clear_other_paragraphs=True)
    sel_lamp_nomor = tabel_lampiran.rows[1].cells[2]  # biarkan seperti template
    # (tidak diubah)

    # 5) Tabel 2 (Daftar Peserta): auto-skalasi + isi
    tabel_peserta = doc.tables[TABLE_PESERTA]
    baris_header = 1
    baris_data_saat_ini = len(tabel_peserta.rows) - baris_header
    baris_data_dibutuhkan = len(peserta)

    if baris_data_dibutuhkan < baris_data_saat_ini:
        _remove_table_rows(tabel_peserta, baris_data_saat_ini - baris_data_dibutuhkan)
    elif baris_data_dibutuhkan > baris_data_saat_ini:
        last_row = tabel_peserta.rows[-1]
        for _ in range(baris_data_dibutuhkan - baris_data_saat_ini):
            _add_row_like(tabel_peserta, last_row)

    # 6) Isi tiap baris peserta
    #    Header (row 0) juga dipaksa Arial 10
    for cell in tabel_peserta.rows[0].cells:
        _apply_cell_font(cell)
    for i, p in enumerate(peserta, start=1):
        _fill_peserta_row(tabel_peserta.rows[i], i, p["nama"], p["nik"], p["divisi"])

    # 7) Update judul daftar peserta (P35) — placeholder 'Judul' bold
    if P_JUDUL_LAMPIRAN < len(doc.paragraphs):
        p_judul_lampiran = doc.paragraphs[P_JUDUL_LAMPIRAN]
        _set_field_bold_placeholder(p_judul_lampiran, "Judul", judul)

    doc.save(output_path)
    return output_path, output_filename
