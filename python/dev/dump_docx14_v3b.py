import sys
sys.path.insert(0, r"D:\Freelaces\Server\htdocs\adk\python")
from docx import Document
from lxml import etree

W  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

def emu_to_cm(emu):
    return round(emu / 360000, 2)

def show(path, label, start=38, end=58):
    print(f"\n{'='*70}")
    print(f"  {label}")
    print('='*70)
    doc = Document(path)
    body_els = list(doc.element.body)
    pg = 1

    for i, el in enumerate(body_els):
        tag = el.tag.split('}')[-1]
        pPr = el.find(f'{{{W}}}pPr') if tag == 'p' else None
        has_sect  = pPr is not None and pPr.find(f'{{{W}}}sectPr') is not None
        has_pgbr  = tag == 'p' and any(br.get(f'{{{W}}}type') == 'page' for br in el.iter(f'{{{W}}}br'))
        if has_sect or has_pgbr:
            pg_before = pg
            pg += 1
        else:
            pg_before = pg

        if start <= i <= end:
            if tag != 'p':
                print(f"  [{i:3}] <{tag}>")
                continue
            txt = ''.join(t.text or '' for t in el.iter(f'{{{W}}}t')).strip()[:50]
            anchors = el.findall(f'.//{{{WP}}}anchor')
            inlines = el.findall(f'.//{{{WP}}}inline')
            img_info = []
            for anc in anchors:
                ext = anc.find(f'{{{WP}}}extent')
                w  = emu_to_cm(int(ext.get('cx', 0))) if ext is not None else 0
                h  = emu_to_cm(int(ext.get('cy', 0))) if ext is not None else 0
                img_info.append(f"ANCHOR({w}x{h}cm)")
            for inl in inlines:
                img_info.append("INLINE")

            sect_detail = ''
            if has_sect and pPr is not None:
                sectPr = pPr.find(f'{{{W}}}sectPr')
                if sectPr is not None:
                    pgNum = sectPr.find(f'{{{W}}}pgNumType')
                    titlePg = sectPr.find(f'{{{W}}}titlePg')
                    if pgNum is not None:
                        start2 = pgNum.get(f'{{{W}}}start', '-')
                        fmt   = pgNum.get(f'{{{W}}}fmt', '-')
                        sect_detail += f" pgNum(start={start2},fmt={fmt})"
                    if titlePg is not None:
                        sect_detail += " titlePg"
                    # Count children
                    n_children = len(list(sectPr))
                    sect_detail += f" [{n_children}attrs]"

            flags = []
            if has_sect: flags.append('SECT')
            if has_pgbr: flags.append('PGBR')
            if img_info: flags += img_info
            flag_str = ' [' + ', '.join(flags) + ']' if flags else ''
            marker = '  <-- BOUNDARY' if (has_sect or has_pgbr) else ''
            disp = repr(txt) if txt else '(empty)'
            print(f"  [{i:3}] pg{pg_before}  {disp}{flag_str}{sect_detail}{marker}")

    # Final body sectPr
    body = doc.element.body
    final_sect = body.find(f'{{{W}}}sectPr')
    if final_sect is not None:
        print(f"\n  [body sectPr] children: {len(list(final_sect))}")

show(
    r"D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\hasil\Docx 14_c1_v3.docx",
    "v3 Docx 14 [38-58]"
)
