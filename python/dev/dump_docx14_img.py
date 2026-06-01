import sys
sys.path.insert(0, r"D:\Freelaces\Server\htdocs\adk\python")
from docx import Document
from lxml import etree

W  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A  = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC= "http://schemas.openxmlformats.org/drawingml/2006/picture"

def emu_to_cm(emu):
    return round(emu / 914400 * 2.54, 2)

def show_anchor_detail(path, label):
    print(f"\n{'='*70}")
    print(f"  {label}")
    print('='*70)
    doc = Document(path)
    body_els = list(doc.element.body)
    pg = 1

    for i, el in enumerate(body_els):
        tag = el.tag.split('}')[-1]
        if tag != 'p':
            continue
        pPr = el.find(f'{{{W}}}pPr')
        has_sect = pPr is not None and pPr.find(f'{{{W}}}sectPr') is not None
        has_pgbr = any(br.get(f'{{{W}}}type') == 'page' for br in el.iter(f'{{{W}}}br'))
        txt  = ''.join(t.text or '' for t in el.iter(f'{{{W}}}t')).strip()[:50]

        # Cek anchor/inline
        anchors = el.findall(f'.//{{{WP}}}anchor')
        inlines = el.findall(f'.//{{{WP}}}inline')

        if anchors or inlines:
            for anc in anchors:
                posH = anc.find(f'{{{WP}}}positionH')
                posV = anc.find(f'{{{WP}}}positionV')
                ext  = anc.find(f'{{{WP}}}extent')  # wp:extent, bukan a:ext
                relFromH = posH.get('relativeFrom') if posH is not None else '?'
                relFromV = posV.get('relativeFrom') if posV is not None else '?'
                offsetH  = posH.find(f'{{{WP}}}posOffset') if posH is not None else None
                offsetV  = posV.find(f'{{{WP}}}posOffset') if posV is not None else None
                alignH   = posH.find(f'{{{WP}}}align') if posH is not None else None
                alignV   = posV.find(f'{{{WP}}}align') if posV is not None else None

                h_val = (alignH.text if alignH is not None else
                         (f"{emu_to_cm(int(offsetH.text))} cm" if offsetH is not None and offsetH.text else '?'))
                v_val = (alignV.text if alignV is not None else
                         (f"{emu_to_cm(int(offsetV.text))} cm" if offsetV is not None and offsetV.text else '?'))

                w_emu = int(ext.get('cx', 0)) if ext is not None else 0
                h_emu = int(ext.get('cy', 0)) if ext is not None else 0

                distT = int(anc.get('distT', 0))
                distB = int(anc.get('distB', 0))
                behindDoc = anc.get(f'behindDoc', '0')
                wrap_el = (anc.find(f'{{{WP}}}wrapNone') or
                           anc.find(f'{{{WP}}}wrapSquare') or
                           anc.find(f'{{{WP}}}wrapThrough') or
                           anc.find(f'{{{WP}}}wrapTight') or
                           anc.find(f'{{{WP}}}wrapTopAndBottom'))
                wrap_type = wrap_el.tag.split('}')[-1] if wrap_el is not None else 'unknown'

                print(f"  [{i:3}] pg{pg}  ANCHOR:")
                print(f"         wrap={wrap_type}, behindDoc={behindDoc}")
                print(f"         posH: relativeFrom={relFromH}, value={h_val}")
                print(f"         posV: relativeFrom={relFromV}, value={v_val}")
                print(f"         size: {emu_to_cm(w_emu)} x {emu_to_cm(h_emu)} cm")
                print(f"         distT={emu_to_cm(distT)} cm, distB={emu_to_cm(distB)} cm")
                # Raw XML anchor (ringkas)
                raw = etree.tostring(anc, pretty_print=False).decode()[:400]
                print(f"         raw: {raw}")

            for inl in inlines:
                ext = inl.find(f'.//{{{A}}}ext')
                w_emu = int(ext.get('cx', 0)) if ext is not None else 0
                h_emu = int(ext.get('cy', 0)) if ext is not None else 0
                print(f"  [{i:3}] pg{pg}  INLINE: size {emu_to_cm(w_emu)} x {emu_to_cm(h_emu)} cm")

        if has_sect or has_pgbr:
            flags = []
            if has_sect: flags.append('SECT')
            if has_pgbr: flags.append('PGBR')
            flag_str = '[' + ','.join(flags) + ']'
            print(f"  [{i:3}] pg{pg}  {repr(txt) if txt else '(empty)'}  {flag_str} <-- BOUNDARY")
            pg += 1

        if i > 60:
            print("  ... (truncated)")
            break

    # Juga tampilkan body sectPr
    body = doc.element.body
    final_sect = body.find(f'{{{W}}}sectPr')
    if final_sect is not None:
        pgSz = final_sect.find(f'{{{W}}}pgSz')
        pgMar = final_sect.find(f'{{{W}}}pgMar')
        print(f"\n  [body sectPr] page size: {emu_to_cm(int(pgSz.get(f'{{{W}}}w', 0)))} x {emu_to_cm(int(pgSz.get(f'{{{W}}}h', 0)))} cm" if pgSz is not None else "")

# Dump ORIGINAL
show_anchor_detail(
    r"D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\Docx 14.docx",
    "ORIGINAL Docx 14"
)

# Dump HASIL
show_anchor_detail(
    r"D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\hasil\Docx 14_c1_v2.docx",
    "HASIL Docx 14_c1_v2"
)
