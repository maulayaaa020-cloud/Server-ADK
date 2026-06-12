"""Verifikasi font di TOC2/TOC3 paragraf Docx 2 Hasil vs File Benar."""
import sys; sys.stdout.reconfigure(encoding='utf-8')
import zipfile
from lxml import etree

ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W  = f"{{{ns}}}"
SEP = "=" * 70

def get_first_text_run_font(p):
    """Cari run pertama yang punya teks, cek fontnya (recursive ke hyperlink)."""
    for r in p.iter(f"{W}r"):
        t_el = r.find(f"{W}t")
        if t_el is not None and (t_el.text or "").strip():
            rPr = r.find(f"{W}rPr")
            if rPr is None:
                return "rPr=None"
            fonts = rPr.find(f"{W}rFonts")
            sz    = rPr.find(f"{W}sz")
            font_info = "None"
            if fonts is not None:
                fa = fonts.get(f"{W}ascii")
                ft = fonts.get(f"{W}asciiTheme")
                font_info = fa or ft or "None"
            sz_info = None
            if sz is not None:
                v = sz.get(f"{W}val")
                sz_info = f"{int(v)/2:.0f}pt" if v else None
            return f"font={font_info!r}, sz={sz_info!r}"
    return "no text run"

def check_file(path, label):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    root = etree.fromstring(xml)
    body = root.find(f"{W}body")

    # Ambil semua TOC2/TOC3 paragraf (bisa dari body langsung atau dari SDT)
    all_paras = []
    # Cek SDT dulu
    for child in list(body):
        if child.tag == f"{W}sdt":
            sdtContent = child.find(f"{W}sdtContent")
            if sdtContent is not None:
                for p in sdtContent.findall(f".//{W}p"):
                    pPr = p.find(f"{W}pPr")
                    if pPr is not None:
                        ps = pPr.find(f"{W}pStyle")
                        if ps is not None and ps.get(f"{W}val") in ("TOC2","TOC3"):
                            all_paras.append(p)
            break

    # Kalau tidak ada SDT, ambil dari body langsung
    if not all_paras:
        for child in list(body):
            pPr = child.find(f"{W}pPr")
            if pPr is not None:
                ps = pPr.find(f"{W}pStyle")
                if ps is not None and ps.get(f"{W}val") in ("TOC2","TOC3"):
                    all_paras.append(child)

    print(f"\n{SEP}")
    print(f" {label}: {len(all_paras)} TOC2/TOC3 paragraf")
    print(SEP)
    for p in all_paras[:8]:
        txt = "".join((t.text or "") for t in p.iter(f"{W}t")).strip()[:40]
        fi = get_first_text_run_font(p)
        print(f"  {txt!r:45} | {fi}")
    if len(all_paras) > 8:
        print(f"  ... (+{len(all_paras)-8} lagi)")

check_file(r"D:\Freelaces\Test Dafis\File Benar\Docx 2.docx", "FILE BENAR Docx 2")
check_file(r"D:\Freelaces\Test Dafis\Hasil\Docx 2.docx", "FILE HASIL Docx 2")
