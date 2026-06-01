from docx import Document
from docx.oxml.ns import qn

def inspect_layout(path, label):
    doc = Document(path)
    print(f'\n=== {label} ===')
    for i, sec in enumerate(doc.sections):
        sp = sec._sectPr
        _t = sp.find(qn('w:type'))
        _type = _t.get(qn('w:val')) if _t is not None else 'nextPage'
        pgSz  = sp.find(qn('w:pgSz'))
        pgMar = sp.find(qn('w:pgMar'))
        titlePg  = sp.find(qn('w:titlePg')) is not None
        evenOdd  = sp.find(qn('w:evenAndOddHeaders')) is not None

        W  = qn('w:w');  H  = qn('w:h')
        T  = qn('w:top'); B = qn('w:bottom'); L = qn('w:left'); R = qn('w:right')
        OR = qn('w:orient')

        sz_str  = f"w={pgSz.get(W)} h={pgSz.get(H)} orient={pgSz.get(OR,'portrait')}" if pgSz is not None else 'NONE'
        mar_str = f"top={pgMar.get(T)} bot={pgMar.get(B)} left={pgMar.get(L)} right={pgMar.get(R)}" if pgMar is not None else 'NONE'

        RI = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
        frefs = [(fr.get(qn('w:type')), fr.get(RI)) for fr in sp.findall(qn('w:footerReference'))]
        hrefs = [(hr.get(qn('w:type')), hr.get(RI)) for hr in sp.findall(qn('w:headerReference'))]

        print(f'  sec[{i}] type={_type} titlePg={titlePg} evenOdd={evenOdd}')
        print(f'         sz  : {sz_str}')
        print(f'         mar : {mar_str}')
        print(f'         hdr : {hrefs}')
        print(f'         ftr : {frefs}')

SRC = r'D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\Docx 15.docx'
V5  = r'D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 1 dimulai dari ii\hasil\Docx 15_c1_v5.docx'
inspect_layout(SRC, 'ORIGINAL')
inspect_layout(V5,  'V5 OUTPUT')
