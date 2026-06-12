"""Bandingkan konten ZIP antara ORIG, BENAR, HASIL untuk Docx 1."""
import zipfile, hashlib

ORIG  = r"D:\Freelaces\Test Dafis\Docx 1.docx"
BENAR = r"D:\Freelaces\Test Dafis\File Benar\Docx 1.docx"
HASIL = r"D:\Freelaces\Test Dafis\Hasil\Docx 1.docx"

def zip_manifest(path):
    with zipfile.ZipFile(path) as z:
        return {
            name: (info.file_size, hashlib.md5(z.read(name)).hexdigest())
            for name in z.namelist()
            for info in [z.getinfo(name)]
        }

orig  = zip_manifest(ORIG)
benar = zip_manifest(BENAR)
hasil = zip_manifest(HASIL)

all_keys = sorted(set(list(orig) + list(benar) + list(hasil)))

print(f"{'FILE':<40} {'ORIG':>10} {'BENAR':>10} {'HASIL':>10}  STATUS")
print("-"*90)
for k in all_keys:
    os_ = orig.get(k)
    bs_ = benar.get(k)
    hs_ = hasil.get(k)

    def sz(x): return f"{x[0]:,}" if x else "-"
    def same(a, b): return a and b and a[1] == b[1]

    status = []
    if not same(os_, bs_):
        status.append("O!=B")
    if not same(os_, hs_):
        status.append("O!=H")
    if not same(bs_, hs_):
        status.append("B!=H")

    marker = "  " if not status else "! "
    print(f"{marker}{k:<40} {sz(os_):>10} {sz(bs_):>10} {sz(hs_):>10}  {' '.join(status)}")

# Fokus ke word/document.xml — cari perbedaan di luar area DAFTAR ISI
print("\n\n=== PERBEDAAN word/document.xml ORIG vs BENAR (di luar area TOC) ===")
from lxml import etree

ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"

def get_paras(path):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    tree = etree.fromstring(xml)
    return tree.find(f"{W}body").findall(f"{W}p")

paras_o = get_paras(ORIG)
paras_b = get_paras(BENAR)

print(f"ORIG: {len(paras_o)} paras | BENAR: {len(paras_b)} paras")
# Bandingkan per-paragraf
diffs = 0
for i in range(max(len(paras_o), len(paras_b))):
    xo = etree.tostring(paras_o[i]).decode() if i < len(paras_o) else ""
    xb = etree.tostring(paras_b[i]).decode() if i < len(paras_b) else ""
    if xo != xb:
        to = "".join(t.text or "" for t in paras_o[i].iter(f"{W}t")).strip() if i < len(paras_o) else ""
        tb = "".join(t.text or "" for t in paras_b[i].iter(f"{W}t")).strip() if i < len(paras_b) else ""
        print(f"  [{i}] ORIG={repr(to[:40])} vs BENAR={repr(tb[:40])}")
        diffs += 1
        if diffs >= 20:
            print("  ... (lebih banyak diff, dipotong)")
            break

if diffs == 0:
    print("  document.xml identik di semua paragraf!")
