"""
Post-process: tambahkan explicit font (Times New Roman, 12pt) ke setiap run
di paragraf TOC2/TOC3 agar XML-nya sama persis dengan File Benar.

Bisa dijalankan untuk file apapun — hanya menyentuh TOC2/TOC3 paragraf.
"""
import sys, zipfile, shutil, os
sys.stdout.reconfigure(encoding='utf-8')
from lxml import etree

NS  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W   = f"{{{NS}}}"


def _ensure_rPr(r_el):
    """Ambil atau buat w:rPr di dalam run."""
    rPr = r_el.find(f"{W}rPr")
    if rPr is None:
        rPr = etree.SubElement(r_el, f"{W}rPr")
        r_el.insert(0, rPr)
    return rPr


def _set_run_font(rPr):
    """Tambahkan/update rFonts ascii + hAnsi = Times New Roman."""
    fonts = rPr.find(f"{W}rFonts")
    if fonts is None:
        fonts = etree.Element(f"{W}rFonts")
        rPr.insert(0, fonts)
    fonts.set(f"{W}ascii",  "Times New Roman")
    fonts.set(f"{W}hAnsi",  "Times New Roman")
    # Hapus asciiTheme jika ada (theme font yang menyebabkan Calibri)
    if f"{W}asciiTheme" in fonts.attrib:
        del fonts.attrib[f"{W}asciiTheme"]
    if f"{W}hAnsiTheme" in fonts.attrib:
        del fonts.attrib[f"{W}hAnsiTheme"]


def _set_run_size(rPr, half_pts="24"):
    """Tambahkan/update sz dan szCs."""
    for tag in (f"{W}sz", f"{W}szCs"):
        el = rPr.find(tag)
        if el is None:
            el = etree.SubElement(rPr, tag)
        el.set(f"{W}val", half_pts)


def fix_toc_font_in_docx(path, styles_to_fix=("TOC2", "TOC3")):
    """
    Buka docx, tambahkan explicit TNR 12pt ke semua run di TOC2/TOC3 paragraf,
    simpan kembali ke file yang sama.
    """
    tmp = path + ".tmp"
    shutil.copy2(path, tmp)
    changed = 0

    with zipfile.ZipFile(tmp, "r") as zin:
        names = zin.namelist()
        xmls = {n: zin.read(n) for n in names}

    # Modifikasi document.xml
    xml = xmls["word/document.xml"]
    root = etree.fromstring(xml)
    body = root.find(f"{W}body")

    for child in list(body):
        pPr = child.find(f"{W}pPr")
        if pPr is None:
            continue
        pStyle_el = pPr.find(f"{W}pStyle")
        if pStyle_el is None:
            continue
        style_val = pStyle_el.get(f"{W}val", "")
        if style_val not in styles_to_fix:
            continue

        # Semua run di paragraf ini (termasuk nested di dalam w:hyperlink, dll.)
        for r_el in child.findall(f".//{W}r"):
            rPr = _ensure_rPr(r_el)
            _set_run_font(rPr)
            _set_run_size(rPr, "24")
            changed += 1

    xmls["word/document.xml"] = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )

    # Tulis ulang zip
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, xmls[name])

    os.remove(tmp)
    print(f"  Fixed {changed} runs in {os.path.basename(path)}")
    return changed


if __name__ == "__main__":
    target = r"D:\Freelaces\Test Dafis\Hasil\Docx 2.docx"
    fix_toc_font_in_docx(target)
    print("Selesai.")
