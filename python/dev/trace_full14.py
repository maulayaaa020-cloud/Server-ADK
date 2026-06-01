"""Trace full pipeline untuk Docx 14"""
import os, sys
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(__file__), '..')))
from docx import Document
from utils import DocProcessor, is_roman_start

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
ROOT = os.path.normpath(os.path.join(os.path.dirname(__file__), '..', '..'))
SRC  = os.path.join(ROOT, 'test_files', 'paket3', 'cover 1 dimulai dari ii', 'Docx 14.docx')

doc  = Document(SRC)
proc = DocProcessor(doc, 'Times New Roman', 12)

def has_sectPr(el):
    pPr = el.find(f'{{{W}}}pPr')
    return pPr is not None and pPr.find(f'{{{W}}}sectPr') is not None

def has_anchor(el):
    return el.find(f'.//{{{WP}}}anchor') is not None

def snapshot(label, idxs):
    body = list(doc.element.body)
    print(f"\n  [{label}]")
    for i in idxs:
        if i >= len(body): continue
        el = body[i]
        tag = el.tag.split('}')[-1]
        if tag != 'p':
            print(f"    body[{i}] <{tag}>")
            continue
        txt = ''.join(t.text or '' for t in el.iter(f'{{{W}}}t')).strip()[:30]
        flags = []
        if has_sectPr(el): flags.append('SECT')
        if has_anchor(el): flags.append('ANCHOR')
        disp = repr(txt) if txt else '(empty)'
        print(f"    body[{i}] {disp} {flags}")

print("=== STEP 0: purge_all_headers_footers (seperti main.py) ===")
proc.purge_all_headers_footers()
snapshot("post-purge", range(40, 50))

print("=== STEP 1: scan_zones ===")
rsp, bab = proc.scan_zones()
body = list(doc.element.body)
rsp_i = body.index(rsp)
print(f"  rsp=[{rsp_i}] {repr(DocProcessor._p_text(rsp)[:30])}")
snapshot("pre-advance", range(40, 50))

print("\n=== STEP 2: advance_roman_start ===")
new_rsp, use_exact = DocProcessor.advance_roman_start(doc, rsp, 1)
body2 = list(doc.element.body)
new_rsp_i = body2.index(new_rsp) if new_rsp in body2 else -1
print(f"  new_rsp=[{new_rsp_i}] {repr(DocProcessor._p_text(new_rsp)[:30])}, exact={use_exact}")
snapshot("post-advance", range(40, 50))

# Kita update roman_start_p sesuai main.py logic
if new_rsp is not rsp:
    rsp = new_rsp
_advanced = use_exact

print("\n=== STEP 3: insert_breaks ===")
rsp = proc.insert_breaks(rsp, bab, exact_roman_start=_advanced)
body3 = list(doc.element.body)
rsp_i3 = body3.index(rsp) if rsp in body3 else -1
print(f"  rsp after insert_breaks=[{rsp_i3}] {repr(DocProcessor._p_text(rsp)[:30])}")
snapshot("post-insert_breaks", range(40, 50))

print("\n=== STEP 4: build_section_map ===")
roman_sec, bab_sec_list, n_sections = proc.build_section_map(rsp, bab)
print(f"  roman_sec={roman_sec}, n_sections={n_sections}")

print("\n=== STEP 5: paket3.apply ===")
import paket3
paket3.apply(proc, roman_sec, bab_sec_list, n_sections, 'Ya', dimulai_dari='ii', num_cover=1)
snapshot("post-paket3", range(40, 50))

print("\n=== STEP 6: sanitize_margins + save ===")
try:
    proc.sanitize_margins()
except Exception as e:
    print(f"  sanitize_margins error: {e}")
snapshot("post-sanitize", range(40, 50))

OUT = r"D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\hasil\trace_v3_check.docx"
doc.save(OUT)
print(f"  Saved to {OUT}")

# Reload dan cek
from docx import Document as Doc2
doc2 = Doc2(OUT)
body_saved = list(doc2.element.body)
print(f"\n=== POST-SAVE reload [40-50] ===")
for i in range(40, min(51, len(body_saved))):
    el = body_saved[i]
    tag2 = el.tag.split('}')[-1]
    if tag2 != 'p':
        print(f"  [{i}] <{tag2}>")
        continue
    txt2 = ''.join(t.text or '' for t in el.iter(f'{{{W}}}t')).strip()[:30]
    pPr2 = el.find(f'{{{W}}}pPr')
    sp2 = pPr2 is not None and pPr2.find(f'{{{W}}}sectPr') is not None
    anc2 = el.find(f'.//{{{WP}}}anchor') is not None
    flags2 = []
    if sp2: flags2.append('SECT')
    if anc2: flags2.append('ANCHOR')
    disp2 = repr(txt2) if txt2 else '(empty)'
    print(f"  [{i}] {disp2} {flags2}")

print("\n=== DONE: final body[40-50] ===")
snapshot("final", range(40, 50))
