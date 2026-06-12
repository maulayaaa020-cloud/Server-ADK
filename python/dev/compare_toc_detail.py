"""
Bandingkan konten TOC (entries, font, size, bold, spacing, indent) antara
File Benar asli dan HASIL secara detail.
"""
import zipfile
from lxml import etree

ns  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W   = f"{{{ns}}}"

BENAR = r"D:\Freelaces\Test Dafis\File Benar\Docx 1.docx"
HASIL = r"D:\Freelaces\Test Dafis\Hasil\Docx 1.docx"


def qn(tag):
    return f"{W}{tag.split(':')[-1]}"


def get_val(el, attr):
    if el is None:
        return None
    v = el.get(f"{W}{attr}")
    return v


def parse_rPr(rPr):
    if rPr is None:
        return {}
    info = {}
    fonts = rPr.find(f"{W}rFonts")
    if fonts is not None:
        info['font_ascii']  = fonts.get(f"{W}ascii")
        info['font_hAnsi']  = fonts.get(f"{W}hAnsi")
        info['font_theme']  = fonts.get(f"{W}asciiTheme")
    sz = rPr.find(f"{W}sz")
    if sz is not None:
        v = sz.get(f"{W}val")
        info['sz_half'] = v
        info['sz_pt']   = f"{int(v)/2:.1f}pt" if v else None
    b = rPr.find(f"{W}b")
    if b is not None:
        info['bold'] = b.get(f"{W}val", 'true')
    else:
        info['bold'] = None
    return info


def parse_pPr(pPr):
    if pPr is None:
        return {}
    info = {}
    pStyle = pPr.find(f"{W}pStyle")
    info['pStyle'] = pStyle.get(f"{W}val") if pStyle is not None else None
    spacing = pPr.find(f"{W}spacing")
    if spacing is not None:
        info['spAfter']  = spacing.get(f"{W}after")
        info['spBefore'] = spacing.get(f"{W}before")
        info['line']     = spacing.get(f"{W}line")
        info['lineRule'] = spacing.get(f"{W}lineRule")
    ind = pPr.find(f"{W}ind")
    if ind is not None:
        info['indLeft']    = ind.get(f"{W}left")
        info['indHanging'] = ind.get(f"{W}hanging")
    tabs = pPr.find(f"{W}tabs")
    if tabs is not None:
        tab_list = []
        for tab in tabs.findall(f"{W}tab"):
            tab_list.append({
                'val':    tab.get(f"{W}val"),
                'pos':    tab.get(f"{W}pos"),
                'leader': tab.get(f"{W}leader"),
            })
        info['tabs'] = tab_list
    return info


def extract_toc_entries(path, label):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    root = etree.fromstring(xml)
    body = root.find(f"{W}body")

    # Cari SDT daftar isi
    sdt = None
    for child in list(body):
        if child.tag == f"{W}sdt":
            sdt = child
            break

    if sdt is None:
        print(f"[{label}] Tidak ada SDT ditemukan!")
        return []

    sdtContent = sdt.find(f"{W}sdtContent")
    if sdtContent is None:
        print(f"[{label}] SDT tidak punya sdtContent!")
        return []

    paras = sdtContent.findall(f".//{W}p")
    print(f"\n{'='*70}")
    print(f" {label}: {len(paras)} paragraf di sdtContent")
    print(f"{'='*70}")

    entries = []
    for i, p in enumerate(paras):
        # Kumpulkan teks
        txt = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()

        pPr  = p.find(f"{W}pPr")
        pp   = parse_pPr(pPr)

        # Ambil rPr dari run pertama yang ada teks
        first_rPr = None
        for r in p.findall(f"{W}r"):
            t_el = r.find(f"{W}t")
            if t_el is not None and (t_el.text or "").strip():
                first_rPr = r.find(f"{W}rPr")
                break
        # Fallback: rPr di pPr
        if first_rPr is None and pPr is not None:
            first_rPr = pPr.find(f"{W}rPr")

        rp = parse_rPr(first_rPr)

        # Cek apakah ada fldChar (instrText = field, bukan teks biasa)
        instrs = p.findall(f".//{W}instrText")
        is_field = bool(instrs)
        instr_txt = (instrs[0].text or "")[:30] if instrs else ""

        entry = {
            'i': i,
            'txt': txt[:60],
            'pStyle': pp.get('pStyle'),
            'spAfter': pp.get('spAfter'),
            'spBefore': pp.get('spBefore'),
            'line': pp.get('line'),
            'lineRule': pp.get('lineRule'),
            'indLeft': pp.get('indLeft'),
            'indHanging': pp.get('indHanging'),
            'tabs': pp.get('tabs'),
            'font_ascii': rp.get('font_ascii'),
            'font_theme': rp.get('font_theme'),
            'sz_pt': rp.get('sz_pt'),
            'sz_half': rp.get('sz_half'),
            'bold': rp.get('bold'),
            'is_field': is_field,
            'instr': instr_txt,
        }
        entries.append(entry)

        # Print ringkas
        field_tag = f" [FIELD:{instr_txt}]" if is_field else ""
        print(f"  [{i:2}] style={str(pp.get('pStyle')):12} | "
              f"font={str(rp.get('font_ascii') or rp.get('font_theme')):18} | "
              f"sz={str(rp.get('sz_pt')):8} | "
              f"bold={str(rp.get('bold')):6} | "
              f"spAfter={str(pp.get('spAfter')):5} | "
              f"indL={str(pp.get('indLeft')):6} | "
              f"hang={str(pp.get('indHanging')):6} | "
              f"{repr(txt[:35]):40}{field_tag}")

    return entries


def compare(e_benar, e_hasil):
    print(f"\n{'='*70}")
    print(" PERBANDINGAN: BENAR vs HASIL (entry per entry)")
    print(f"{'='*70}")

    # Cocokkan berdasarkan teks (fuzzy: cukup prefix 20 char)
    def key(e):
        return e['txt'][:20].strip()

    map_benar = {key(e): e for e in e_benar if not e['is_field']}
    map_hasil = {key(e): e for e in e_hasil if not e['is_field']}

    # Entries di BENAR tapi tidak di HASIL
    only_benar = set(map_benar) - set(map_hasil)
    only_hasil = set(map_hasil) - set(map_benar)
    common     = set(map_benar) & set(map_hasil)

    if only_benar:
        print(f"\n  [HANYA DI BENAR - tidak ada di HASIL]:")
        for k in sorted(only_benar):
            print(f"    '{k}'")

    if only_hasil:
        print(f"\n  [HANYA DI HASIL - tidak ada di BENAR]:")
        for k in sorted(only_hasil):
            print(f"    '{k}'")

    print(f"\n  [COMMON: {len(common)} entries — bandingkan detail]")
    FIELDS = ['pStyle', 'font_ascii', 'sz_pt', 'bold', 'spAfter', 'spBefore',
              'line', 'lineRule', 'indLeft', 'indHanging']

    diffs = []
    for k in sorted(common):
        eb = map_benar[k]
        eh = map_hasil[k]
        diff_fields = [(f, eb[f], eh[f]) for f in FIELDS if eb.get(f) != eh.get(f)]
        if diff_fields:
            diffs.append((k, diff_fields))

    if diffs:
        print(f"\n  PERBEDAAN DITEMUKAN di {len(diffs)} entry:")
        for k, df in diffs:
            print(f"\n    Entry: '{k}'")
            for fname, vb, vh in df:
                print(f"      {fname:12}: BENAR={str(vb):15} vs HASIL={str(vh)}")
    else:
        print("\n  Semua common entry: format SAMA ✅")

    # Tab stops comparison (khusus sample H1, H2, H3)
    print(f"\n  TAB STOPS (sample dari tiap level):")
    for label_list, entries in [("BENAR", e_benar), ("HASIL", e_hasil)]:
        styles_seen = set()
        for e in entries:
            s = e['pStyle']
            if s and s not in styles_seen and e['tabs']:
                styles_seen.add(s)
                print(f"    [{label_list}] {s}: {e['tabs']}")
            if len(styles_seen) >= 4:
                break


e_benar = extract_toc_entries(BENAR, "FILE BENAR (asli)")
e_hasil = extract_toc_entries(HASIL, "FILE HASIL")
compare(e_benar, e_hasil)
