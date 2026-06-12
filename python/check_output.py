import zipfile
from lxml import etree

path = r"C:\Users\farizal\Downloads\Test Dafis - Output.docx"
ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

with zipfile.ZipFile(path) as z:
    xml = z.read("word/document.xml")

tree = etree.fromstring(xml)
body = tree.find(f"{{{ns}}}body")
paras = body.findall(f".//{{{ns}}}p")

print("=== CHECK TOC OUTPUT ===")
in_toc = False
count = 0
for i, p in enumerate(paras):
    runs = p.findall(f".//{{{ns}}}t")
    text = "".join(r.text or "" for r in runs).strip()

    pPr = p.find(f"{{{ns}}}pPr")
    style = ""
    if pPr is not None:
        pStyle = pPr.find(f"{{{ns}}}pStyle")
        if pStyle is not None:
            style = pStyle.get(f"{{{ns}}}val", "")

    if "DAFTAR ISI" in text.upper() and not in_toc:
        in_toc = True
        print(f"[{i}] HEADER: '{text}' (style={style})")
        continue

    if in_toc:
        # check section break
        sect = p.find(f".//{{{ns}}}sectPr")
        if sect is not None:
            print(f"  [{i}] <<SECTION BREAK>> style={style} text='{text[:40]}'")

        if text:
            # check italic in any run
            rPrs = p.findall(f".//{{{ns}}}rPr")
            italic_runs = []
            for rPr in rPrs:
                i_el = rPr.find(f"{{{ns}}}i")
                if i_el is not None:
                    val = i_el.get(f"{{{ns}}}val", "1")
                    italic_runs.append(val)

            # check indent
            ind_info = ""
            if pPr is not None:
                ind = pPr.find(f"{{{ns}}}ind")
                if ind is not None:
                    left = ind.get(f"{{{ns}}}left", "?")
                    hang = ind.get(f"{{{ns}}}hanging", "0")
                    ind_info = f" [left={left},hang={hang}]"

            italic_flag = "ITALIC!" if any(v not in ("0", "false") for v in italic_runs) else ""
            print(f"  [{i}] {style:15} {italic_flag:8} '{text[:65]}'{ind_info}")
            count += 1
            if count > 45:
                print("  ... truncated")
                break
        else:
            if sect is None:
                pass  # skip empty non-section paras
