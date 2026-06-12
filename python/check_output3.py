import zipfile
from lxml import etree

path = r"C:\Users\farizal\Downloads\Test Dafis - Output6.docx"
ns  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

with zipfile.ZipFile(path) as z:
    doc_xml    = z.read("word/document.xml")
    styles_xml = z.read("word/styles.xml")

dtree = etree.fromstring(doc_xml)
body  = dtree.find(f"{{{ns}}}body")
sectPr = body.find(f"{{{ns}}}sectPr")
if sectPr is None:
    for p in reversed(list(body)):
        sp = p.find(f"{{{ns}}}pPr/{{{ns}}}sectPr")
        if sp is None:
            sp = p.find(f"{{{ns}}}sectPr")
        if sp is not None:
            sectPr = sp
            break

def tw(el, attr):
    return int(el.get(f"{{{ns}}}{attr}", 0)) if el is not None else 0

pgSz  = sectPr.find(f"{{{ns}}}pgSz")  if sectPr is not None else None
pgMar = sectPr.find(f"{{{ns}}}pgMar") if sectPr is not None else None
w      = tw(pgSz,  'w')
left   = tw(pgMar, 'left')
right  = tw(pgMar, 'right')
text_w = w - left - right

print("=== Dimensi dokumen ===")
print(f"  page_w : {w}  ({w/567:.1f} cm)")
print(f"  left   : {left}  ({left/567:.1f} cm)")
print(f"  right  : {right}  ({right/567:.1f} cm)")
print(f"  text_w : {text_w}  ({text_w/567:.1f} cm)  <- seharusnya posisi right tab")

stree = etree.fromstring(styles_xml)
print("\n=== Tab stop kanan di TOC styles ===")
for s in stree.findall(f".//{{{ns}}}style"):
    sid = s.get(f"{{{ns}}}styleId", "")
    if not (sid.startswith("TOC") or sid.startswith("toc")):
        continue
    name_el = s.find(f"{{{ns}}}name")
    name = name_el.get(f"{{{ns}}}val", "") if name_el is not None else ""
    for tab in s.findall(f".//{{{ns}}}tab"):
        if tab.get(f"{{{ns}}}val") == "right":
            pos = tab.get(f"{{{ns}}}pos", "?")
            print(f"  {name:15} right_tab={pos} ({int(pos)/567:.1f} cm)")
    pPr = s.find(f"{{{ns}}}pPr")
    if pPr is not None:
        ind = pPr.find(f"{{{ns}}}ind")
        jc  = pPr.find(f"{{{ns}}}jc")
        if ind is not None:
            print(f"  {name:15} indent left={ind.get(f'{{{ns}}}left','0')} hang={ind.get(f'{{{ns}}}hanging','0')}")
        if jc is not None:
            print(f"  {name:15} align={jc.get(f'{{{ns}}}val')}")
