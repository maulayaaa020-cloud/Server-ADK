import sys, os
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(__file__), '..')))
from docx import Document
from docx.oxml.ns import qn
from utils import DocProcessor

W  = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
RI = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'

SRC = r'D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\Docx 15.docx'

def show_sections(doc, label):
    print(f'\n  [{label}] {len(list(doc.sections))} sections')
    for i, sec in enumerate(doc.sections):
        sp = sec._sectPr
        t = sp.find(qn('w:type'))
        tval = t.get(qn('w:val')) if t is not None else 'nextPage'
        frefs = [(fr.get(qn('w:type')), fr.get(RI)) for fr in sp.findall(qn('w:footerReference'))]
        hrefs = [(hr.get(qn('w:type')), hr.get(RI)) for hr in sp.findall(qn('w:headerReference'))]
        # Find which paragraph contains this sectPr (if embedded)
        para_txt = ''
        body = list(doc.element.body)
        for el in body:
            if not el.tag.endswith('}p'): continue
            pPr = el.find(f'{{{W}}}pPr')
            if pPr is not None and pPr.find(f'{{{W}}}sectPr') is sp:
                para_txt = ''.join(t2.text or '' for t2 in el.iter(f'{{{W}}}t')).strip()[:25]
                break
        loc = f' @ "{para_txt}"' if para_txt else ' @ [body]'
        print(f'    sec[{i}] {tval:12} hdr={hrefs} ftr={frefs}{loc}')

doc  = Document(SRC)
proc = DocProcessor(doc, 'Times New Roman', 12)

print('=== STEP 0: initial ===')
show_sections(doc, 'initial')

print('\n=== STEP 1: purge_all_headers_footers ===')
proc.purge_all_headers_footers()
show_sections(doc, 'after purge')

print('\n=== STEP 2: scan_zones ===')
rsp, bab = proc.scan_zones()
body = list(doc.element.body)
print(f'  roman_start_p: [{body.index(rsp)}] "{DocProcessor._p_text(rsp)[:30]}"')
print(f'  bab_p_list: {len(bab)} items')
for b in bab:
    print(f'    [{body.index(b)}] "{DocProcessor._p_text(b)[:25]}"')

print('\n=== STEP 3: advance_roman_start ===')
new_rsp, use_exact = DocProcessor.advance_roman_start(doc, rsp, 1)
body = list(doc.element.body)
print(f'  new_rsp=[{body.index(new_rsp)}] "{DocProcessor._p_text(new_rsp)[:30]}" exact={use_exact}')
if new_rsp is not rsp: rsp = new_rsp

print('\n=== STEP 4: insert_breaks ===')
rsp = proc.insert_breaks(rsp, bab, exact_roman_start=use_exact)
body = list(doc.element.body)
print(f'  rsp after=[{body.index(rsp) if rsp in body else -1}] "{DocProcessor._p_text(rsp)[:30]}"')
show_sections(doc, 'after insert_breaks')

# Show body[25-35] around DAFTAR ISI
print('\n  Body[25-38] after insert_breaks:')
for i in range(25, 39):
    if i >= len(body): break
    el = body[i]
    tag = el.tag.split('}')[-1]
    if tag != 'p':
        print(f'    [{i}] <{tag}>')
        continue
    txt = ''.join(t.text or '' for t in el.iter(f'{{{W}}}t')).strip()[:30]
    pPr = el.find(f'{{{W}}}pPr')
    sp = pPr.find(f'{{{W}}}sectPr') if pPr is not None else None
    flags = []
    if sp is not None:
        t = sp.find(f'{{{W}}}type')
        flags.append(f'SECT({t.get(qn("w:val")) if t is not None else "nextPage"})')
    print(f'    [{i}] {"(empty)" if not txt else repr(txt)} {flags}')

print('\n=== STEP 5: build_section_map ===')
roman_sec, bab_sec_list, n_sections = proc.build_section_map(rsp, bab)
print(f'  roman_sec={roman_sec}, bab_sec_list={bab_sec_list}, n_sections={n_sections}')
