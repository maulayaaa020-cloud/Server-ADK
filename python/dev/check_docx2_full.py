"""Cek semua TOC entries Docx 2 Hasil (non-SDT) dan bandingkan font/bold dengan File Benar."""
import sys; sys.stdout.reconfigure(encoding='utf-8')
import zipfile
from lxml import etree

ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"
SEP = "=" * 70

def parse_rPr(rPr):
    if rPr is None: return {}
    info = {}
    fonts = rPr.find(f"{W}rFonts")
    if fonts is not None:
        info["font_ascii"] = fonts.get(f"{W}ascii")
        info["font_theme"] = fonts.get(f"{W}asciiTheme")
    sz = rPr.find(f"{W}sz")
    if sz is not None:
        v = sz.get(f"{W}val")
        info["sz_pt"] = f"{int(v)/2:.1f}pt" if v else None
    b = rPr.find(f"{W}b")
    info["bold"] = b.get(f"{W}val", "true") if b is not None else None
    return info

def parse_pPr(pPr):
    if pPr is None: return {}
    info = {}
    pStyle = pPr.find(f"{W}pStyle")
    info["pStyle"] = pStyle.get(f"{W}val") if pStyle is not None else None
    return info

def extract_from_body_by_style(path, label):
    """Extract TOC entries dengan TOC1/TOC2/TOC3 style langsung dari body."""
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    root = etree.fromstring(xml)
    body = root.find(f"{W}body")
    children = list(body)
    print(f"\n{SEP}")
    print(f" {label}")
    print(SEP)
    entries = []
    for i, child in enumerate(children):
        pPr = child.find(f"{W}pPr")
        pp = parse_pPr(pPr)
        pstyle = (pp.get("pStyle") or "").lower().strip()
        if pstyle not in ("toc1","toc2","toc3","toc 1","toc 2","toc 3","tocheading","toc heading"):
            continue
        txt = "".join((t.text or "") for t in child.findall(f".//{W}t")).strip()
        first_rPr = None
        for r in child.findall(f"{W}r"):
            t_el = r.find(f"{W}t")
            if t_el is not None and (t_el.text or "").strip():
                first_rPr = r.find(f"{W}rPr")
                break
        if first_rPr is None and pPr is not None:
            first_rPr = pPr.find(f"{W}rPr")
        rp = parse_rPr(first_rPr)
        instrs = child.findall(f".//{W}instrText")
        is_field = bool(instrs)
        font_disp = rp.get("font_ascii") or rp.get("font_theme") or "None"
        print(f"  [{i:3}] style={pp.get('pStyle'):12} | font={font_disp:20} | "
              f"sz={str(rp.get('sz_pt')):8} | bold={str(rp.get('bold')):6} | {txt[:45]!r}")
        entries.append({"i": i, "txt": txt[:60], "pStyle": pp.get("pStyle"),
                        "is_field": is_field, **rp})
    return entries

def extract_sdt_toc(path, label):
    """Extract TOC entries dari format SDT."""
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
        print(f"[{label}] Tidak ada SDT — coba extract dari body style")
        return extract_from_body_by_style(path, label)
    sdtContent = sdt.find(f"{W}sdtContent")
    paras = list(sdtContent.findall(f".//{W}p")) if sdtContent else []
    print(f"\n{SEP}")
    print(f" {label} (SDT): {len(paras)} paragraf")
    print(SEP)
    entries = []
    for i, p in enumerate(paras):
        txt = "".join((t.text or "") for t in p.iter(f"{W}t")).strip()
        pPr = p.find(f"{W}pPr")
        pp = parse_pPr(pPr)
        pstyle = pp.get("pStyle") or ""
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
        font_disp = rp.get("font_ascii") or rp.get("font_theme") or "None"
        print(f"  [{i:2}] style={pstyle:12} | font={font_disp:20} | "
              f"sz={str(rp.get('sz_pt')):8} | bold={str(rp.get('bold')):6} | {txt[:45]!r}")
        if txt:
            entries.append({"i": i, "txt": txt[:60], "pStyle": pstyle,
                            "is_field": is_field, **rp})
    return entries

e_benar = extract_sdt_toc(
    r"D:\Freelaces\Test Dafis\File Benar\Docx 2.docx",
    "FILE BENAR Docx 2"
)
e_hasil = extract_from_body_by_style(
    r"D:\Freelaces\Test Dafis\Hasil\Docx 2.docx",
    "FILE HASIL Docx 2"
)

# Compare text content
print(f"\n{SEP}\n PERBANDINGAN ENTRIES (nama saja, tanpa nomor halaman)\n{SEP}")

def strip_page(txt):
    """Hapus nomor halaman di akhir (digit/roman)."""
    import re
    return re.sub(r'[\divxlcXIVLC]+$', '', txt).strip()

def key(e):
    return strip_page(e["txt"])[:20].strip()

entries_b = [e for e in e_benar if not e.get("is_field", False) and e["txt"].strip()]
entries_h = [e for e in e_hasil if not e.get("is_field", False) and e["txt"].strip()]

map_b = {key(e): e for e in entries_b}
map_h = {key(e): e for e in entries_h}
only_b = set(map_b) - set(map_h)
only_h = set(map_h) - set(map_b)
common = set(map_b) & set(map_h)

print(f"  BENAR: {len(map_b)} entries, HASIL: {len(map_h)} entries")
if only_b:
    print("\n  [HANYA DI BENAR]:")
    for k in sorted(only_b): print(f"    {k!r}")
if only_h:
    print("\n  [HANYA DI HASIL]:")
    for k in sorted(only_h): print(f"    {k!r}")

FIELDS = ["pStyle", "font_ascii", "sz_pt", "bold"]
diffs = []
for k in sorted(common):
    eb, eh = map_b[k], map_h[k]
    df = [(f, eb.get(f), eh.get(f)) for f in FIELDS if eb.get(f) != eh.get(f)]
    if df: diffs.append((k, df))
if diffs:
    print(f"\n  PERBEDAAN FORMAT di {len(diffs)} entry:")
    for k, df in diffs:
        print(f"\n    {k!r}")
        for fname, vb, vh in df:
            print(f"      {fname}: BENAR={str(vb)!r:15} | HASIL={str(vh)!r}")
else:
    print(f"\n  Format sama persis di {len(common)} common entries! OK")
