import sys
sys.path.insert(0, r"D:\Freelaces\Server\htdocs\adk\python")
from docx import Document

W  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

def emu_to_cm(e): return round(e/360000, 2)

def dump(path, label):
    doc = Document(path)
    body = list(doc.element.body)
    print(f"\n=== {label} ===")
    pg = 1
    for i, el in enumerate(body):
        tag = el.tag.split('}')[-1]
        if tag != 'p': continue
        pPr = el.find(f'{{{W}}}pPr')
        has_sect = pPr is not None and pPr.find(f'{{{W}}}sectPr') is not None
        has_pgbr = any(br.get(f'{{{W}}}type') == 'page' for br in el.iter(f'{{{W}}}br'))
        txt = ''.join(t.text or '' for t in el.iter(f'{{{W}}}t')).strip()[:40]

        anchors = el.findall(f'.//{{{WP}}}anchor')
        inlines = el.findall(f'.//{{{WP}}}inline')
        img_info = []
        for anc in anchors:
            ext = anc.find(f'{{{WP}}}extent')
            w = emu_to_cm(int(ext.get('cx',0))) if ext is not None else 0
            h = emu_to_cm(int(ext.get('cy',0))) if ext is not None else 0
            img_info.append(f"ANCHOR({w}x{h}cm)")
        for inl in inlines:
            ext = inl.find(f'{{{WP}}}extent')
            w = emu_to_cm(int(ext.get('cx',0))) if ext is not None else 0
            img_info.append(f"INLINE({w}cm)")

        sect_d = ''
        if has_sect and pPr is not None:
            sp = pPr.find(f'{{{W}}}sectPr')
            pn = sp.find(f'{{{W}}}pgNumType') if sp is not None else None
            tpg = sp.find(f'{{{W}}}titlePg') if sp is not None else None
            if pn is not None:
                sect_d += f" pgNum(s={pn.get(f'{{{W}}}start','-')},f={pn.get(f'{{{W}}}fmt','-')})"
            if tpg is not None:
                sect_d += " titlePg"

        if 38 <= i <= 52:
            flags = []
            if has_sect: flags.append('SECT')
            if has_pgbr: flags.append('PGBR')
            flags += img_info
            flag_str = ' ['+', '.join(flags)+']' if flags else ''
            disp = repr(txt) if txt else '(empty)'
            mark = '  <--' if (has_sect or has_pgbr) else ''
            print(f"  [{i:3}] pg{pg} {disp}{flag_str}{sect_d}{mark}")

        if has_sect or has_pgbr:
            pg += 1
        if i > 55: break

dump(
    r"D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\hasil\Docx 14_c1_v4.docx",
    "v4 — argumen benar"
)
