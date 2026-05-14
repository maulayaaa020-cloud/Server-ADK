"""
main.py — Entry point pemrosesan dokumen ADK.
Usage: python main.py <input.docx> <output.docx> [paket] [font] [size] [hidden_cover] [posisi]

Paket:
  paket1 — Full Angka, posisi bebas
  paket2 — Romawi + Angka, posisi bebas seragam
  paket3 — Romawi + Angka, posisi tetap (Populer Skripsi)  [default]
"""
import sys
import re
import json
import os
import zipfile
import tempfile

# Pastikan folder python/ ada di sys.path agar import utils/paket* bisa berjalan
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import paket1
import paket2
import paket3
from utils import DocProcessor


def _fail(code, message):
    print(json.dumps({"status": "error", "code": code, "message": message}))
    sys.exit(1)


# ── LibreOffice artifact cleaner ──────────────────────────────────────────────
_W14_NS         = 'http://schemas.microsoft.com/office/word/2010/wordml'
_MC_NS          = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
_LO_NS_PREFIXES = ('urn:org:documentfoundation', 'com.sun.star', 'urn:openoffice')

_CLEAN_XML_FILES = frozenset([
    'word/document.xml', 'word/settings.xml', 'word/styles.xml',
    'word/endnotes.xml', 'word/footnotes.xml',
])


def _strip_lo_artifacts(input_path):
    """
    Preprocess DOCX: strip LibreOffice-specific XML yang menyebabkan
    peringatan 'unreadable content' di Microsoft Word.
    Returns path ke temp file yang sudah bersih (caller harus unlink).
    """
    from lxml import etree

    def _is_lo_tag(tag):
        if not isinstance(tag, str) or not tag.startswith('{'):
            return False
        ns = tag[1:tag.index('}')]
        return any(ns.startswith(p) for p in _LO_NS_PREFIXES)

    def _clean_xml(data, is_settings=False):
        try:
            root = etree.fromstring(data)
        except Exception:
            return data  # biarkan python-docx menangani jika parse gagal

        # 1. Hapus w14:conflictMode dari settings.xml — paling sering menyebabkan warning
        if is_settings:
            for elem in list(root.iter('{%s}conflictMode' % _W14_NS)):
                p = elem.getparent()
                if p is not None:
                    p.remove(elem)

        # 2. Hapus elemen dengan namespace LibreOffice-specific (bottom-up)
        for elem in reversed(list(root.iter())):
            if _is_lo_tag(elem.tag):
                p = elem.getparent()
                if p is not None:
                    p.remove(elem)

        # 3. Unwrap mc:AlternateContent → ambil isi mc:Fallback (lebih kompatibel)
        AC = '{%s}AlternateContent' % _MC_NS
        FB = '{%s}Fallback' % _MC_NS
        CH = '{%s}Choice' % _MC_NS
        changed = True
        while changed:
            changed = False
            for ac in list(root.iter(AC)):
                parent = ac.getparent()
                if parent is None:
                    continue
                idx = list(parent).index(ac)
                src = ac.find(FB) or ac.find(CH)
                if src is not None:
                    for i, child in enumerate(list(src)):
                        parent.insert(idx + i, child)
                parent.remove(ac)
                changed = True
                break

        return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    tmp_fd, tmp_path = tempfile.mkstemp(suffix='.docx')
    os.close(tmp_fd)

    with zipfile.ZipFile(input_path, 'r') as zin, \
         zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            raw = zin.read(item.filename)
            fn  = item.filename
            if fn.endswith('.xml') and (
                fn in _CLEAN_XML_FILES or
                fn.startswith('word/header') or
                fn.startswith('word/footer')
            ):
                raw = _clean_xml(raw, is_settings=(fn == 'word/settings.xml'))
            zout.writestr(item, raw)

    return tmp_path


def validate_file(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == '.doc':
        _fail("FORMAT_NOT_SUPPORTED",
              "Format .doc tidak didukung. Buka di Microsoft Word lalu simpan ulang sebagai .docx.")

    try:
        size = os.path.getsize(path)
    except OSError:
        _fail("FILE_READ_ERROR", "Tidak dapat membaca ukuran file.")

    if size > 30 * 1024 * 1024:
        _fail("FILE_TOO_LARGE",
              f"Ukuran file ({size // (1024*1024)} MB) melebihi batas maksimal 30 MB.")

    try:
        with zipfile.ZipFile(path, 'r') as z:
            names = z.namelist()

            if 'word/document.xml' not in names:
                _fail("INVALID_DOCX", "File bukan dokumen Word (.docx) yang valid.")

            # Tolak file bermacro (docm yang diganti ekstensi ke docx)
            if 'word/vbaProject.bin' in names:
                _fail("MACRO_DETECTED",
                      "File mengandung macro VBA dan tidak dapat diproses. "
                      "Simpan ulang sebagai .docx biasa (bukan .docm).")

            # Zip bomb: total ukuran isi tidak boleh melebihi 200 MB
            total_uncompressed = sum(info.file_size for info in z.infolist())
            if total_uncompressed > 200 * 1024 * 1024:
                _fail("FILE_TOO_LARGE",
                      "Konten file terlalu besar untuk diproses (melebihi 200 MB tidak terkompresi).")

    except zipfile.BadZipFile:
        _fail("INVALID_DOCX", "File rusak atau bukan format .docx yang valid.")


def main():
    if len(sys.argv) < 3:
        _fail("MISSING_ARGUMENTS",
              "Usage: python main.py <input_file> <output_file> "
              "[paket] [font] [size] [hidden_cover] [posisi]")

    input_file  = sys.argv[1]
    output_file = sys.argv[2]
    paket       = sys.argv[3] if len(sys.argv) > 3 else 'paket3'
    font_arg    = sys.argv[4] if len(sys.argv) > 4 else 'Times New Roman'
    size_arg    = sys.argv[5] if len(sys.argv) > 5 else '12 pt'
    hidden_cov  = sys.argv[6] if len(sys.argv) > 6 else 'Ya'
    posisi      = sys.argv[7] if len(sys.argv) > 7 else 'Tengah Bawah'

    m            = re.search(r'\d+', size_arg)
    font_size_pt = int(m.group()) if m else 12

    validate_file(input_file)

    # ── Bersihkan artifact LibreOffice, lalu buka dokumen ────────────────────
    _temp_cleaned = None
    try:
        _temp_cleaned = _strip_lo_artifacts(input_file)
        open_target   = _temp_cleaned
    except Exception:
        open_target   = input_file  # fallback ke file asli jika strip gagal

    try:
        from docx import Document
        doc = Document(open_target)
    except Exception as e:
        _fail("FILE_READ_ERROR", f"Gagal membuka file: {e}")
    finally:
        if _temp_cleaned and _temp_cleaned != input_file:
            try:
                os.unlink(_temp_cleaned)
            except OSError:
                pass

    # ── Proses ───────────────────────────────────────────
    try:
        proc = DocProcessor(doc, font_arg, font_size_pt)
        proc.purge_all_headers_footers()

        detected_bab_texts = []
        if paket == 'paket1':
            # Paket 1: tidak butuh deteksi zona
            paket1.apply(proc, hidden_cov, posisi)

        else:
            # Paket 2 & 3: butuh deteksi zona → Phase 1-3 dulu
            roman_start_p, bab_p_list = proc.scan_zones()
            detected_bab_texts = [
                DocProcessor._p_text(p)[:60] for p in bab_p_list
            ]
            roman_start_p             = proc.insert_breaks(roman_start_p, bab_p_list)
            roman_sec, bab_sec_list, n_sections = proc.build_section_map(
                roman_start_p, bab_p_list
            )

            if paket == 'paket2':
                paket2.apply(proc, roman_sec, bab_sec_list, n_sections, hidden_cov, posisi)
            else:
                paket3.apply(proc, roman_sec, bab_sec_list, n_sections, hidden_cov)

    except Exception as e:
        _fail("PROCESSING_ERROR", f"Gagal memproses dokumen: {e}")

    # ── Simpan ───────────────────────────────────────────
    try:
        doc.save(output_file)
    except Exception as e:
        _fail("FILE_SAVE_ERROR", f"Gagal menyimpan file: {e}")

    # ── Output ringkasan (dibaca PHP) ─────────────────────
    out_paras  = list(doc.paragraphs)
    out_breaks = [i for i, p in enumerate(out_paras) if DocProcessor._has_sectPr(p._p)]
    out_bounds = [0] + [bi + 1 for bi in out_breaks]

    sections_info = []
    for sec_idx, start in enumerate(out_bounds):
        end        = out_bounds[sec_idx + 1] if sec_idx + 1 < len(out_bounds) else len(out_paras)
        first_text = ""
        for i in range(start, min(start + 10, end)):
            t = out_paras[i].text.strip()
            if t:
                first_text = t[:50]
                break
        sections_info.append({"index": sec_idx, "first_content": first_text})

    print(json.dumps({
        "status":         "success",
        "paket":          paket,
        "total_sections": len(doc.sections),
        "detected_bab":   detected_bab_texts if paket != 'paket1' else [],
        "sections":       sections_info
    }))


if __name__ == '__main__':
    main()
