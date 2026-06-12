"""Lihat isi SDT (Structured Document Tag) yang menyimpan TOC di BENAR Docx 1."""
import zipfile
from lxml import etree

BENAR = r"D:\Freelaces\Test Dafis\File Benar\Docx 1.docx"
HASIL = r"D:\Freelaces\Test Dafis\Hasil\Docx 1.docx"

ns  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W   = f"{{{ns}}}"

def inspect_sdt(path, label):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    root = etree.fromstring(xml)
    body = root.find(f"{W}body")
    children = list(body)

    sdts = [c for c in children if c.tag == f"{W}sdt"]
    print(f"\n=== {label}: {len(sdts)} SDT ditemukan ===")
    for si, sdt in enumerate(sdts):
        sdtPr = sdt.find(f"{W}sdtPr")
        sdtContent = sdt.find(f"{W}sdtContent")

        # Cek sdtPr
        if sdtPr is not None:
            alias   = sdtPr.find(f"{W}alias")
            tag     = sdtPr.find(f"{W}tag")
            docPart = sdtPr.find(f".//{W}docPartGallery")
            print(f"  SDT [{si}] sdtPr:")
            if alias   is not None: print(f"    alias  = {alias.get(f'{W}val','')}")
            if tag     is not None: print(f"    tag    = {tag.get(f'{W}val','')}")
            if docPart is not None: print(f"    gallery= {docPart.get(f'{W}val','')}")
            # Tampilkan semua child sdtPr
            for ch in sdtPr:
                tname = ch.tag.split('}')[1] if '}' in ch.tag else ch.tag
                print(f"    {tname}: {etree.tostring(ch).decode()[:100]}")

        # Lihat isi sdtContent
        if sdtContent is not None:
            content_paras = sdtContent.findall(f".//{W}p")
            print(f"  SDT [{si}] sdtContent: {len(content_paras)} paragraf")
            for j, p in enumerate(content_paras[:6]):
                txt  = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
                flds = p.findall(f".//{W}fldChar")
                instrs = p.findall(f".//{W}instrText")
                pPr  = p.find(f"{W}pPr")
                sn   = pPr.find(f"{W}pStyle") if pPr is not None else None
                style = sn.get(f"{W}val","") if sn is not None else ""
                flags = []
                if flds:   flags.append(f"fld:{','.join(f.get(f'{W}fldCharType','?') for f in flds)}")
                if instrs: flags.append(f"instr:{(instrs[0].text or '')[:25]}")
                print(f"    [{j}] {style:10} | {repr(txt[:45]):50} | {' '.join(flags)}")
            if len(content_paras) > 6:
                print(f"    ... ({len(content_paras)-6} lagi)")

        raw = etree.tostring(sdt, pretty_print=True).decode()
        print(f"  SDT [{si}] total XML: {len(raw)} chars\n")

inspect_sdt(BENAR, "BENAR")
inspect_sdt(HASIL, "HASIL")

# Tampilkan raw sdtPr dari BENAR SDT pertama
print("=== RAW sdtPr BENAR SDT [0] ===")
with zipfile.ZipFile(BENAR) as z:
    xml = z.read("word/document.xml")
root = etree.fromstring(xml)
body = root.find(f"{W}body")
sdt = next(c for c in body if c.tag == f"{W}sdt")
sdtPr = sdt.find(f"{W}sdtPr")
print(etree.tostring(sdtPr, pretty_print=True).decode())
