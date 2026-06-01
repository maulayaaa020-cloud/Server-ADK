"""Bandingkan trace_v3_check.docx vs Docx 14_c1_v3.docx"""
import sys
sys.path.insert(0, r"D:\Freelaces\Server\htdocs\adk\python")
from docx import Document

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

def has_sectPr(el):
    pPr = el.find(f'{{{W}}}pPr')
    return pPr is not None and pPr.find(f'{{{W}}}sectPr') is not None

def has_anchor(el):
    return el.find(f'.//{{{WP}}}anchor') is not None

def dump(path, label):
    doc = Document(path)
    body = list(doc.element.body)
    print(f"\n=== {label} ===")
    for i in range(40, min(47, len(body))):
        el = body[i]
        tag = el.tag.split('}')[-1]
        if tag != 'p':
            print(f"  [{i}] <{tag}>"); continue
        txt = ''.join(t.text or '' for t in el.iter(f'{{{W}}}t')).strip()[:20]
        flags = []
        if has_sectPr(el): flags.append('SECT')
        if has_anchor(el): flags.append('ANCHOR')
        disp = repr(txt) if txt else '(empty)'
        print(f"  [{i}] {disp} {flags}")

dump(
    r"D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\hasil\trace_v3_check.docx",
    "trace_v3_check.docx"
)
dump(
    r"D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\hasil\Docx 14_c1_v3.docx",
    "Docx 14_c1_v3.docx (main.py)"
)
