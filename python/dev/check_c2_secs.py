import sys, os
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(__file__), '..')))
from docx import Document
from docx.oxml.ns import qn

hasil = r'D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 2 dimulai dari iii\hasil'
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

for n in range(1, 19):
    path = os.path.join(hasil, f'Docx {n}_c2.docx')
    if not os.path.exists(path):
        continue
    doc  = Document(path)
    secs = list(doc.sections)
    info = []
    for i, sec in enumerate(secs[:5]):
        pn    = sec._sectPr.find(qn('w:pgNumType'))
        fmt   = pn.get(qn('w:fmt'))   if pn is not None else None
        start = pn.get(qn('w:start')) if pn is not None else None
        ftr   = len(sec._sectPr.findall(qn('w:footerReference')))
        info.append(f'[{i}]fmt={fmt} start={start} ftr={ftr}')
    print(f'Docx {n:2d}: {" | ".join(info[:4])}')
