"""Debug advance_roman_start untuk Docx 14"""
import sys
sys.path.insert(0, r"D:\Freelaces\Server\htdocs\adk\python")
from docx import Document
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

path = r"D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\Docx 14.docx"
doc = Document(path)
body_els = list(doc.element.body)

def has_sectPr(el):
    pPr = el.find(f'{{{W}}}pPr')
    return pPr is not None and pPr.find(f'{{{W}}}sectPr') is not None

def has_anchor(el):
    WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    return el.find(f'.//{{{WP}}}anchor') is not None

# Show paragraphs 40-50 with sectPr/anchor flags
print("=== Original body [40-50] ===")
for i in range(40, min(51, len(body_els))):
    el = body_els[i]
    tag = el.tag.split('}')[-1]
    if tag != 'p':
        print(f"  [{i}] <{tag}>")
        continue
    txt = ''.join(t.text or '' for t in el.iter(f'{{{W}}}t')).strip()[:40]
    flags = []
    if has_sectPr(el): flags.append('SECT')
    if has_anchor(el): flags.append('ANCHOR')
    disp = repr(txt) if txt else '(empty)'
    print(f"  [{i}] {disp} {flags}")

# Now simulate the advance logic
print("\n=== Simulating advance_roman_start fix ===")
# Find [43] = image+sectPr
cand_idx = 44  # candidate is [44] based on earlier analysis
_prev_idx = cand_idx - 1  # = 43
_prev = body_els[_prev_idx]

print(f"_prev = [{_prev_idx}], has_anchor={has_anchor(_prev)}, has_sectPr={has_sectPr(_prev)}")

_pPr_prev = _prev.find(f'{{{W}}}pPr')
print(f"_pPr_prev = {_pPr_prev}")

if _pPr_prev is not None:
    _sectPr_mv = _pPr_prev.find(f'{{{W}}}sectPr')
    print(f"_sectPr_mv = {_sectPr_mv}")

    if _sectPr_mv is not None:
        _pPr_prev.remove(_sectPr_mv)
        print("Removed sectPr from [43]")

        # Search backward for empty paragraph
        _body_now = list(doc.element.body)
        _prev_bi = next((i for i, e in enumerate(_body_now) if e is _prev), -1)
        print(f"_prev in body_now at index: {_prev_bi}")

        for _k in range(_prev_bi - 1, max(-1, _prev_bi - 10), -1):
            _ek = _body_now[_k]
            tag_k = _ek.tag.split('}')[-1]
            txt_k = ''.join(t.text or '' for t in _ek.iter(f'{{{W}}}t')).strip()[:20]
            has_c = has_anchor(_ek)
            has_s = has_sectPr(_ek)
            print(f"  Checking [{_k}]: tag={tag_k}, txt={repr(txt_k)}, anchor={has_c}, sect={has_s}")

            WP2 = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            _is_empty = (
                tag_k == 'p' and
                not txt_k and
                not has_c and
                not has_s
            )
            print(f"    → is_empty_candidate: {_is_empty}")
            if _is_empty:
                print(f"    → FOUND empty target at [{_k}]")
                break
