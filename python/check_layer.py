import sys, re
sys.path.insert(0, r'D:\Freelaces\Server\htdocs\adk\python')
from docx import Document
from docx.oxml.ns import qn

doc = Document(r'C:\Users\farizal\Downloads\Test Dafis.docx')

def _get_num_level(para):
    pPr = para._p.find(qn('w:pPr'))
    if pPr is None: return None
    numPr = pPr.find(qn('w:numPr'))
    if numPr is None: return None
    numId_el = numPr.find(qn('w:numId'))
    if numId_el is not None:
        try:
            if int(numId_el.get(qn('w:val'), 1)) == 0: return None
        except: pass
    ilvl_el = numPr.find(qn('w:ilvl'))
    if ilvl_el is None: return None
    try: return int(ilvl_el.get(qn('w:val'), 0))
    except: return None

def _get_outline_xml(para):
    pPr = para._p.find(qn('w:pPr'))
    if pPr is None: return None
    el = pPr.find(qn('w:outlineLvl'))
    if el is None: return None
    val = el.get(qn('w:val'))
    try: lvl = int(val); return lvl + 1 if lvl < 9 else None
    except: return None

def _get_outline_style(para):
    style = para.style
    while style:
        pPr = style.element.find(qn('w:pPr'))
        if pPr is not None:
            el = pPr.find(qn('w:outlineLvl'))
            if el is not None:
                try:
                    lvl = int(el.get(qn('w:val'), 9))
                    if lvl < 9: return lvl + 1
                except: pass
        style = style.base_style
    return None

LIST_STYLES = {'List Paragraph', 'List Bullet', 'List Number',
               'List Bullet 2', 'List Bullet 3', 'List Number 2', 'List Number 3'}

print("=== Paragraf yang berpotensi salah deteksi (panjang > 80, non-BAB) ===")
for i, para in enumerate(doc.paragraphs):
    text = para.text.strip()
    if not text or len(text) > 150 or len(text) <= 20: continue
    if re.match(r'^\s*(BAB|\d+\.)', text, re.IGNORECASE): continue

    style_name = para.style.name if para.style else ''
    num_lvl    = _get_num_level(para)
    ol_xml     = _get_outline_xml(para)
    ol_style   = _get_outline_style(para)

    detected = None
    reason   = ''

    if num_lvl is not None and 1 <= num_lvl <= 2 and style_name not in LIST_STYLES:
        detected = num_lvl + 1
        reason   = f'step4-numPr(ilvl={num_lvl})'
    elif ol_xml is not None:
        detected = ol_xml
        reason   = f'step5-outlineLvl_xml={ol_xml}'
    elif ol_style is not None and len(text) <= 80:
        detected = ol_style
        reason   = f'step5b-style(<=80)'
    elif style_name.startswith('Heading ') and len(text) <= 80:
        detected = int(style_name.split()[-1])
        reason   = f'step7-heading_style'

    if detected is not None and len(text) > 80:
        print(f'[{i}] H{detected} via {reason} | style={style_name} | len={len(text)}')
        print(f'     "{text[:80]}..."')
