"""Cek isi TOC Docx 2 Hasil (format non-SDT) dan File Benar."""
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

def extract_non_sdt_toc(path, label):
    """Extract TOC entries dari format non-SDT (paragraf dengan TOC style)."""
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    root = etree.fromstring(xml)
    body = root.find(f"{W}body")
    children = list(body)

    # Cari paragraf TOC heading
    in_toc = False
    entries = []
    print(f"\n{SEP}")
    print(f" {label} (non-SDT format)")
    print(SEP)

    for i, child in enumerate(children):
        tag = child.tag.split("}")[-1]
        txt = "".join((t.text or "") for t in child.findall(f".//{W}t")).strip()
        pPr = child.find(f"{W}pPr")
        pp = parse_pPr(pPr)
        pstyle = pp.get("pStyle", "") or ""

        # Masuk ke TOC saat ketemu TOCHeading
        if pstyle.lower() in ("toc heading", "tocheading", "daftar isi"):
            in_toc = True
            print(f"  [{i}] TOCHeading: {txt!r}")
            continue

        if in_toc:
            # Keluar dari TOC saat paragraf tanpa TOC style dan ada konten non-TOC
            if not pstyle.lower().startswith("toc") and not pstyle.lower().startswith("daftar isi"):
                # Cek apakah ada instrText TOC
                instrs = child.findall(f".//{W}instrText")
                if not instrs:
                    if txt and len(txt) > 2:
                        print(f"  [END-TOC at {i}] style={pstyle!r} txt={txt[:40]!r}")
                        break
                    continue

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
            print(f"  [{i:3}] style={pstyle:12} | font={font_disp:20} | "
                  f"sz={str(rp.get('sz_pt')):8} | bold={str(rp.get('bold')):6} | "
                  f"{txt[:45]!r}")
            if not is_field and txt:
                entries.append({"i": i, "txt": txt[:60], "pStyle": pstyle, **rp})

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
        print(f"[{label}] Tidak ada SDT!")
        return []

    sdtContent = sdt.find(f"{W}sdtContent")
    if sdtContent is None:
        return []
    paras = list(sdtContent.findall(f".//{W}p"))
    print(f"\n{SEP}")
    print(f" {label} (SDT format): {len(paras)} paragraf")
    print(SEP)
    entries = []
    for i, p in enumerate(paras):
        txt = "".join((t.text or "") for t in p.iter(f"{W}t")).strip()
        pPr = p.find(f"{W}pPr")
        pp = parse_pPr(pPr)
        pstyle = pp.get("pStyle", "") or ""
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
        if not is_field and txt:
            entries.append({"i": i, "txt": txt[:60], "pStyle": pstyle, **rp})
    return entries

# File Benar Docx 2 — SDT format
e_benar = extract_sdt_toc(
    r"D:\Freelaces\Test Dafis\File Benar\Docx 2.docx",
    "FILE BENAR Docx 2"
)

# File Hasil Docx 2 — non-SDT format
e_hasil = extract_non_sdt_toc(
    r"D:\Freelaces\Test Dafis\Hasil\Docx 2.docx",
    "FILE HASIL Docx 2"
)

# Compare
print(f"\n{SEP}\n PERBANDINGAN ENTRIES\n{SEP}")
def key(e): return e["txt"][:20].strip()
map_b = {key(e): e for e in e_benar}
map_h = {key(e): e for e in e_hasil}
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
