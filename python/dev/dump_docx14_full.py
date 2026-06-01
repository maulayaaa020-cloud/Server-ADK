import sys
sys.path.insert(0, r"D:\Freelaces\Server\htdocs\adk\python")
from docx import Document
from lxml import etree

W  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

def emu_to_cm(emu):
    return round(emu / 360000, 2)  # 360000 EMU = 1 cm

def twip_to_cm(twip):
    return round(twip * 2.54 / 1440, 2)

def show_full(path, label):
    print(f"\n{'='*70}")
    print(f"  {label}")
    print('='*70)
    doc = Document(path)
    body_els = list(doc.element.body)
    pg = 1

    for i, el in enumerate(body_els):
        tag = el.tag.split('}')[-1]
        if tag != 'p':
            print(f"  [{i:3}] <{tag}>")
            continue

        pPr = el.find(f'{{{W}}}pPr')
        has_sect  = pPr is not None and pPr.find(f'{{{W}}}sectPr') is not None
        has_pgbr  = any(br.get(f'{{{W}}}type') == 'page' for br in el.iter(f'{{{W}}}br'))
        txt = ''.join(t.text or '' for t in el.iter(f'{{{W}}}t')).strip()[:60]

        # Detect images
        anchors = el.findall(f'.//{{{WP}}}anchor')
        inlines = el.findall(f'.//{{{WP}}}inline')
        img_flags = []
        for anc in anchors:
            ext = anc.find(f'{{{WP}}}extent')
            w  = emu_to_cm(int(ext.get('cx', 0))) if ext is not None else 0
            h  = emu_to_cm(int(ext.get('cy', 0))) if ext is not None else 0
            posV = anc.find(f'{{{WP}}}positionV')
            relV = posV.get('relativeFrom','?') if posV is not None else '?'
            offV = posV.find(f'{{{WP}}}posOffset') if posV is not None else None
            alignV = posV.find(f'{{{WP}}}align') if posV is not None else None
            vval = (alignV.text if alignV is not None else
                    (f"{emu_to_cm(int(offV.text))}cm" if offV is not None and offV.text else '?'))
            bd = anc.get('behindDoc','0')
            img_flags.append(f"ANCHOR({w}x{h}cm,posV={relV}:{vval},behindDoc={bd})")
        for inl in inlines:
            ext = inl.find(f'{{{WP}}}extent')
            w = emu_to_cm(int(ext.get('cx', 0))) if ext is not None else 0
            h = emu_to_cm(int(ext.get('cy', 0))) if ext is not None else 0
            img_flags.append(f"INLINE({w}x{h}cm)")

        # sectPr details
        sect_detail = ''
        if has_sect and pPr is not None:
            sectPr = pPr.find(f'{{{W}}}sectPr')
            if sectPr is not None:
                pgSz = sectPr.find(f'{{{W}}}pgSz')
                pgNum = sectPr.find(f'{{{W}}}pgNumType')
                titlePg = sectPr.find(f'{{{W}}}titlePg')
                if pgSz is not None:
                    w_t = int(pgSz.get(f'{{{W}}}w', 0))
                    h_t = int(pgSz.get(f'{{{W}}}h', 0))
                    sect_detail += f" pgSz={twip_to_cm(w_t)}x{twip_to_cm(h_t)}cm"
                if pgNum is not None:
                    start = pgNum.get(f'{{{W}}}start', '-')
                    fmt   = pgNum.get(f'{{{W}}}fmt', '-')
                    sect_detail += f" pgNum(start={start},fmt={fmt})"
                if titlePg is not None:
                    sect_detail += " titlePg"

        flags = []
        if has_sect: flags.append('SECT')
        if has_pgbr: flags.append('PGBR')
        if img_flags: flags += img_flags
        flag_str = ' [' + ', '.join(flags) + ']' if flags else ''
        marker = '  <-- BOUNDARY' if (has_sect or has_pgbr) else ''

        disp = repr(txt) if txt else '(empty)'
        print(f"  [{i:3}] pg{pg}  {disp}{flag_str}{sect_detail}{marker}")

        if has_sect or has_pgbr:
            pg += 1

        if i >= 60:
            print("  ... (truncated at 60)")
            break

    # Final body sectPr
    body = doc.element.body
    final_sect = body.find(f'{{{W}}}sectPr')
    if final_sect is not None:
        pgSz = final_sect.find(f'{{{W}}}pgSz')
        pgNum = final_sect.find(f'{{{W}}}pgNumType')
        titlePg = final_sect.find(f'{{{W}}}titlePg')
        if pgSz is not None:
            w_t = int(pgSz.get(f'{{{W}}}w', 0))
            h_t = int(pgSz.get(f'{{{W}}}h', 0))
            print(f"\n  [final sectPr] pgSz={twip_to_cm(w_t)}x{twip_to_cm(h_t)}cm", end='')
        if pgNum is not None:
            print(f" pgNum(start={pgNum.get(f'{{{W}}}start','-')},fmt={pgNum.get(f'{{{W}}}fmt','-')})", end='')
        if titlePg is not None:
            print(" titlePg", end='')
        print()

show_full(
    r"D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\Docx 14.docx",
    "ORIGINAL Docx 14"
)
show_full(
    r"D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\hasil\Docx 14_c1_v2.docx",
    "HASIL Docx 14_c1_v2"
)
