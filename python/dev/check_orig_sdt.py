"""Cek posisi SDT di ORIG Docx 1 menggunakan list(body) bukan findall."""
import zipfile
from lxml import etree

ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"

ORIG  = r"D:\Freelaces\Test Dafis\Docx 1.docx"
HASIL = r"D:\Freelaces\Test Dafis\Hasil\Docx 1.docx"

def scan_body_children(path, label, start=44, count=14):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    root = etree.fromstring(xml)
    body = root.find(f"{W}body")
    children = list(body)

    print(f"\n=== {label}: body children [{start}..{start+count}] ===")
    for i in range(start, min(start+count, len(children))):
        child = children[i]
        tag = child.tag.split('}')[1] if '}' in child.tag else child.tag
        if tag == "p":
            txt = "".join(t.text or "" for t in child.iter(f"{W}t")).strip()
            flds = child.findall(f".//{W}fldChar")
            bms  = child.findall(f".//{W}bookmarkStart")
            sp   = child.find(f".//{W}sectPr")
            flags = []
            if flds: flags.append(f"fld")
            if bms:  flags.append(f"bm:{len(bms)}")
            if sp:   flags.append("sect")
            print(f"  [{i}] <p>  {repr(txt[:45]):48} {' '.join(flags)}")
        elif tag == "sdt":
            sdtContent = child.find(f"{W}sdtContent")
            inner_paras = sdtContent.findall(f".//{W}p") if sdtContent is not None else []
            flds = child.findall(f".//{W}fldChar")
            instrs = child.findall(f".//{W}instrText")
            raw = etree.tostring(child, pretty_print=True).decode()
            print(f"  [{i}] <sdt> ({len(raw)} chars, {len(inner_paras)} inner para, fldChar={len(flds)}, instr={len(instrs)})")
            if sdtContent is not None:
                for j, ip in enumerate(inner_paras[:3]):
                    t2 = "".join(t.text or "" for t in ip.iter(f"{W}t")).strip()
                    f2 = ip.findall(f".//{W}fldChar")
                    i2 = ip.findall(f".//{W}instrText")
                    instr_txt = (i2[0].text or "")[:25] if i2 else ""
                    print(f"           [{j}] {repr(t2[:40]):45} fld={len(f2)} instr={repr(instr_txt)}")
        elif tag == "tbl":
            rows = child.findall(f".//{W}tr")
            print(f"  [{i}] <tbl> ({len(rows)} rows)")
        else:
            print(f"  [{i}] <{tag}>")

scan_body_children(ORIG,  "ORIG")
scan_body_children(HASIL, "HASIL")
