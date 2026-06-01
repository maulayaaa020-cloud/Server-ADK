import sys
sys.path.insert(0, r'd:\Freelaces\Server\htdocs\adk\python')
from docx import Document
from docx.oxml.ns import qn
from utils import DocProcessor

path  = r'd:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\hasil\Docx 19.docx'
doc   = Document(path)
paras = list(doc.paragraphs)

print("=== PAGE BREAK DAN EMPTY PARA DI SEKITARNYA ===\n")
for i, p in enumerate(paras):
    has_pgbr = any(br.get(qn('w:type'), '') == 'page'
                   for br in p._p.findall('.//' + qn('w:br')))
    has_sect = DocProcessor._has_sectPr(p._p)
    if not (has_pgbr or has_sect):
        continue

    # Tampilkan 5 paragraf sebelum dan sesudahnya
    print(f"  --- break di para[{i}] pgbr={has_pgbr} sect={has_sect} "
          f"text={repr(p.text.strip()[:40])} ---")
    for j in range(max(0, i-6), min(len(paras), i+4)):
        pb = any(br.get(qn('w:type'), '') == 'page'
                 for br in paras[j]._p.findall('.//' + qn('w:br')))
        sc = DocProcessor._has_sectPr(paras[j]._p)
        mk = ' <-- BREAK' if (pb or sc) else ''
        t  = repr(paras[j].text.strip()[:50])
        has_content = bool(paras[j].text.strip())
        print(f"  [{j:3d}] content={has_content}  pgbr={pb}  sect={sc}  text={t}{mk}")
    print()
