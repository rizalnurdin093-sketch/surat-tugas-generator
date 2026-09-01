"""Copy footer dari template v3 ke template aktif (ikutin footer template baru)."""
import zipfile, shutil, os, re

SRC = '/tmp/template_footer_v3.docx'          # sumber footer
DST = 'templates_surat/template_surat.docx'   # target footer
BAK = '/tmp/template_sebelum_footer.docx'

shutil.copyfile(DST, BAK)  # backup target

# --- baca footer dari SRC ---
with zipfile.ZipFile(SRC) as z:
    footer_src = z.read('word/footer1.xml')
    # juga catat apa saja yang ada (header/media tidak perlu)

# --- baca seluruh isi DST, lalu ganti footer1.xml & hapus footer rels ---
with zipfile.ZipFile(DST) as z:
    items = {n: z.read(n) for n in z.namelist()}

# ganti footer1.xml
items['word/footer1.xml'] = footer_src
# footer v3 tidak punya gambar/relasi -> hapus footer1.xml.rels supaya tak rujuk gambar
for k in list(items):
    if k == 'word/_rels/footer1.xml.rels':
        del items[k]

tmp = DST + '.tmp'
with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as z:
    for name, data in items.items():
        z.writestr(name, data)
shutil.move(tmp, DST)
print('Footer template diperbarui dari v3')
