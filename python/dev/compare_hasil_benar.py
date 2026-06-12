"""Bandingkan file di Hasil vs File Benar — ringkasan per file."""
import zipfile, os
from lxml import etree

ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"
HASIL = r"D:\Freelaces\Test Dafis\Hasil"
BENAR = r"D:\Freelaces\Test Dafis\File Benar"

def get_toc_area(path, limit=50):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    tree = etree.fromstring(xml)
    body = tree.find(f"{W}body")
    paras = body.findall(f"{W}p")

    di = None
    for i, p in enumerate(paras):
        txt = "".join(t.text or "" for t in p.iter(f"{W}t")).strip().upper()
        if txt == "DAFTAR ISI":
            di = i
            break
    if di is None:
        return None, []

    result = []
    for i in range(di, min(di + limit, len(paras))):
        p = paras[i]
        text = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
        flds   = p.findall(f".//{W}fldChar")
        instrs = p.findall(f".//{W}instrText")
        brs    = p.findall(f".//{W}br")
        pPr    = p.find(f"{W}pPr")
        sn     = pPr.find(f"{W}pStyle") if pPr is not None else None
        style  = sn.get(f"{W}val","") if sn is not None else ""
        flags  = []
        if flds:   flags.append("fld:" + ",".join(f.get(f"{W}fldCharType","?") for f in flds))
        if instrs: flags.append("instr:" + (instrs[0].text or "")[:20].strip())
        if brs:    flags.append("br:" + ",".join(b.get(f"{W}type","soft") for b in brs))
        result.append((style, text, " ".join(flags)))
    return di, result

files = sorted(f for f in os.listdir(HASIL) if f.lower().endswith(".docx"))
print(f"{'FILE':<15} {'STATUS':<10} {'KETERANGAN'}")
print("-"*80)

for fname in files:
    hp = os.path.join(HASIL, fname)
    bp = os.path.join(BENAR, fname)
    if not os.path.exists(bp):
        print(f"{fname:<15} {'NO_REF':<10} File Benar tidak ada")
        continue

    _, rows_h = get_toc_area(hp)
    _, rows_b = get_toc_area(bp)

    # Bandingkan text saja (abaikan nomor halaman yg bisa beda)
    texts_h = [r[1] for r in rows_h]
    texts_b = [r[1] for r in rows_b]
    flags_h = [r[2] for r in rows_h]
    flags_b = [r[2] for r in rows_b]

    # Hitung baris yang berbeda
    maxlen = max(len(texts_h), len(texts_b))
    diff_text  = sum(1 for i in range(maxlen)
                     if (texts_h[i] if i < len(texts_h) else "") !=
                        (texts_b[i] if i < len(texts_b) else ""))
    diff_flags = sum(1 for i in range(maxlen)
                     if (flags_h[i] if i < len(flags_h) else "") !=
                        (flags_b[i] if i < len(flags_b) else ""))
    len_diff = len(texts_h) - len(texts_b)

    if diff_text == 0 and diff_flags == 0:
        print(f"{fname:<15} {'SAMA':<10} -")
    else:
        note = []
        if diff_text > 0:
            note.append(f"teks beda={diff_text}")
        if diff_flags > 0:
            note.append(f"flags beda={diff_flags}")
        if len_diff != 0:
            note.append(f"jml para {'lebih' if len_diff>0 else 'kurang'} {abs(len_diff)}")
        status = "BEDA"
        print(f"{fname:<15} {status:<10} {', '.join(note)}")

print("\nDetail file BEDA:")
for fname in files:
    hp = os.path.join(HASIL, fname)
    bp = os.path.join(BENAR, fname)
    if not os.path.exists(bp): continue

    _, rows_h = get_toc_area(hp)
    _, rows_b = get_toc_area(bp)
    texts_h = [r[1] for r in rows_h]
    texts_b = [r[1] for r in rows_b]
    flags_h = [r[2] for r in rows_h]
    flags_b = [r[2] for r in rows_b]
    maxlen = max(len(texts_h), len(texts_b))
    diffs = [(i,
              texts_h[i] if i < len(texts_h) else "<kosong>",
              texts_b[i] if i < len(texts_b) else "<kosong>",
              flags_h[i] if i < len(flags_h) else "",
              flags_b[i] if i < len(flags_b) else "")
             for i in range(maxlen)
             if (texts_h[i] if i < len(texts_h) else "") != (texts_b[i] if i < len(texts_b) else "")
             or (flags_h[i] if i < len(flags_h) else "") != (flags_b[i] if i < len(flags_b) else "")]
    if not diffs:
        continue

    print(f"\n--- {fname} ---")
    print(f"{'OFF':>4} {'HASIL text':38} | {'BENAR text':38}")
    for i, th, tb, fh, fb in diffs[:15]:
        marker = "!" if th != tb else "~"
        print(f"{i:4} {repr(th[:36]):38} {marker} {repr(tb[:36]):38}")
        if fh != fb:
            print(f"     FH: {fh[:38]:38} | FB: {fb}")
