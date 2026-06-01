"""
inspect_hasil.py — Tampilkan detail section tiap file output di folder hasil/.
Satu file per baris detail: fmt, start, ada/tidak header+footer.
"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(__file__), '..')))

from docx import Document
from docx.oxml.ns import qn

_ROOT = os.path.normpath(os.path.join(os.path.dirname(__file__), '..', '..'))

FOLDERS = [
    {
        'dir'  : os.path.join(_ROOT, 'test_files', 'paket3', 'cover 1 dimulai dari ii', 'hasil'),
        'label': 'cover 1 dimulai dari ii  (dimulai=ii, num_cover=1)',
    },
    {
        'dir'  : os.path.join(_ROOT, 'test_files', 'paket3', 'cover 2 dimulai dari iii', 'hasil'),
        'label': 'cover 2 dimulai dari iii  (dimulai=iii, num_cover=2)',
    },
]

def has_content(part):
    try:
        for p in part.paragraphs:
            if p._p.findall('.//' + qn('w:fldChar')):
                return True
            for r in p.runs:
                if r.text.strip():
                    return True
        return False
    except Exception:
        return False

def inspect_file(path):
    doc = Document(path)
    rows = []
    for i, sec in enumerate(doc.sections):
        pn    = sec._sectPr.find(qn('w:pgNumType'))
        fmt   = pn.get(qn('w:fmt'),   'decimal') if pn is not None else 'decimal'
        start = pn.get(qn('w:start'), '-')       if pn is not None else '-'

        diff  = sec.different_first_page_header_footer

        # regular header/footer
        h_ok = has_content(sec.header)
        f_ok = has_content(sec.footer)

        # first-page header/footer (hanya relevan jika diff=True)
        if diff:
            fph_ok = has_content(sec.first_page_header)
            fpf_ok = has_content(sec.first_page_footer)
            hf_str = f'hdr={h_ok} ftr={f_ok} | 1st_hdr={fph_ok} 1st_ftr={fpf_ok}'
        else:
            hf_str = f'hdr={h_ok} ftr={f_ok}'

        rows.append(f'    sec[{i}]  fmt={fmt:<12} start={start:<4}  {hf_str}')
    return rows

for cfg in FOLDERS:
    hasil_dir = cfg['dir']
    print(f'\n{"="*65}')
    print(f'  {cfg["label"]}')
    print(f'{"="*65}')

    files = sorted(
        f for f in os.listdir(hasil_dir)
        if f.lower().endswith('.docx') and not f.startswith('~')
    )

    for fname in files:
        path = os.path.join(hasil_dir, fname)
        print(f'\n  [{fname}]')
        try:
            rows = inspect_file(path)
            for r in rows:
                print(r)
        except Exception as e:
            print(f'    ERROR: {e}')
