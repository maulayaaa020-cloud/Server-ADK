"""Bandingkan TOC Docx 2: File Benar vs Hasil."""
import sys
sys.stdout.reconfigure(encoding='utf-8')
import zipfile
from lxml import etree

ns  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W   = f"{{{ns}}}"

BENAR = r"D:\Freelaces\Test Dafis\File Benar\Docx 2.docx"
HASIL = r"D:\Freelaces\Test Dafis\Hasil\Docx 2.docx"

SEP = "=" * 70

def parse_rPr(rPr):
    if rPr is None: return {}
    info = {}
    fonts = rPr.find(f"{W}rFonts")
    if fonts is not None:
        info['font_ascii'] = fonts.get(f"{W}ascii")
        info['font_theme'] = fonts.get(f"{W}asciiTheme")
    sz = rPr.find(f"{W}sz")
    if sz is not None:
        v = sz.get(f"{W}val")
        info['sz_pt'] = f"{int(v)/2:.1f}pt" if v else None
    b = rPr.find(f"{W}b")
    info['bold'] = b.get(f"{W}val", 'true') if b is not None else None
    return info

def parse_pPr(pPr):
    if pPr is None: return {}
    info = {}
    pStyle = pPr.find(f"{W}pStyle")
    info['pStyle'] = pStyle.get(f"{W}val") if pStyle is not None else None
    spacing = pPr.find(f"{W}spacing")
    if spacing is not None:
        info['spAfter']  = spacing.get(f"{W}after")
        info['spBefore'] = spacing.get(f"{W}before")
        info['line']     = spacing.get(f"{W}line")
    ind = pPr.find(f"{W}ind")
    if ind is not None:
        info['indLeft']    = ind.get(f"{W}left")
        info['indHanging'] = ind.get(f"{W}hanging")
    return info

def extract(path, label):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    root = etree.fromstring(xml)
    body = root.find(f"{W}body")
    sdt = None
    for child in list(body):
        if child.tag == f"{W}sdt":
            sdt = child
            break
    if sdt is None:
        print(f"[{label}] Tidak ada SDT!")
        return []
    sdtContent = sdt.find(f"{W}sdtContent")
    paras = sdtContent.findall(f".//{W}p") if sdtContent else []
    print(f"\n{SEP}")
    print(f" {label}: {len(paras)} paragraf")
    print(SEP)
    entries = []
    for i, p in enumerate(paras):
        txt = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
        pPr = p.find(f"{W}pPr")
        pp = parse_pPr(pPr)
        first_rPr = None
        for r in p.findall(f"{W}r"):
            t_el = r.find(f"{W}t")
            if t_el is not None and (t_el.text or "").strip():
                first_rPr = r.find(f"{W}rPr")
                break
        if first_rPr is None and pPr is not None:
            first_rPr = pPr.find(f"{W}rPr")
        rp = parse_rPr(first_rPr)
        instrs = p.findall(f".//{W}instrText")
        is_field = bool(instrs)
        entry = {"i": i, "txt": txt[:60], "is_field": is_field}
        entry.update(pp)
        entry.update(rp)
        entries.append(entry)
        font_disp = rp.get("font_ascii") or rp.get("font_theme") or "None"
        print(f"  [{i:2}] style={str(pp.get('pStyle')):12} | font={font_disp:20} | "
              f"sz={str(rp.get('sz_pt')):8} | bold={str(rp.get('bold')):6} | "
              f"{repr(txt[:45])}")
    return entries

e_benar = extract(BENAR, "FILE BENAR Docx 2")
e_hasil = extract(HASIL, "FILE HASIL Docx 2")

def key(e): return e["txt"][:20].strip()
map_b = {key(e): e for e in e_benar if not e["is_field"]}
map_h = {key(e): e for e in e_hasil if not e["is_field"]}
only_b = set(map_b) - set(map_h)
only_h = set(map_h) - set(map_b)
common = set(map_b) & set(map_h)

print(f"\n{SEP}\n PERBEDAAN\n{SEP}")
if only_b:
    print("\n  [HANYA DI BENAR - tidak ada di HASIL]:")
    for k in sorted(only_b): print(f"    {repr(k)}")
if only_h:
    print("\n  [HANYA DI HASIL - tidak ada di BENAR]:")
    for k in sorted(only_h): print(f"    {repr(k)}")

FIELDS = ["pStyle", "font_ascii", "sz_pt", "bold", "spAfter", "spBefore",
          "line", "indLeft", "indHanging"]
diffs = []
for k in sorted(common):
    eb, eh = map_b[k], map_h[k]
    df = [(f, eb.get(f), eh.get(f)) for f in FIELDS if eb.get(f) != eh.get(f)]
    if df:
        diffs.append((k, df))

if diffs:
    print(f"\n  PERBEDAAN di {len(diffs)} entry:")
    for k, df in diffs:
        print(f"\n    Entry: {repr(k)}")
        for fname, vb, vh in df:
            print(f"      {fname:12}: BENAR={str(vb):15} | HASIL={str(vh)}")
else:
    print("\n  Semua sama persis! OK")

print(f"\n  Total: BENAR={len(map_b)}, HASIL={len(map_h)}, "
      f"common={len(common)}, only_benar={len(only_b)}, only_hasil={len(only_h)}")
