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

# Pastikan folder python/ ada di sys.path agar import utils/paket* bisa berjalan
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import paket1
import paket2
import paket3
import paket4
from utils import DocProcessor


def _fail(code, message):
    print(json.dumps({"status": "error", "code": code, "message": message}))
    sys.exit(1)


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
    paket       = sys.argv[3]  if len(sys.argv) > 3  else 'paket3'
    font_arg    = sys.argv[4]  if len(sys.argv) > 4  else 'Times New Roman'
    size_arg    = sys.argv[5]  if len(sys.argv) > 5  else '12 pt'
    hidden_cov  = sys.argv[6]  if len(sys.argv) > 6  else 'Ya'
    posisi      = sys.argv[7]  if len(sys.argv) > 7  else 'Tengah Bawah'
    # Paket 4 – custom per-zona
    pos_bab     = sys.argv[8]  if len(sys.argv) > 8  else 'Tengah Bawah'
    pos_isi_bab = sys.argv[9]  if len(sys.argv) > 9  else 'Kanan Atas'
    dimulai     = sys.argv[10] if len(sys.argv) > 10 else 'i'
    semb_dafus  = sys.argv[11] if len(sys.argv) > 11 else 'Tidak'
    semb_lamprn = sys.argv[12] if len(sys.argv) > 12 else 'Tidak'
    num_cover   = int(sys.argv[13]) if len(sys.argv) > 13 else 1

    m            = re.search(r'\d+', size_arg)
    font_size_pt = int(m.group()) if m else 12

    validate_file(input_file)

    # ── Buka dokumen ─────────────────────────────────────
    try:
        from docx import Document
        doc = Document(input_file)
    except Exception as e:
        _fail("FILE_READ_ERROR", f"Gagal membuka file: {e}")

    # ── Proses ───────────────────────────────────────────
    try:
        proc = DocProcessor(doc, font_arg, font_size_pt)
        proc.purge_all_headers_footers()

        detected_bab_texts = []
        if paket == 'paket1':
            # Paket 1: tidak butuh deteksi zona
            paket1.apply(proc, hidden_cov, posisi, dimulai_dari=dimulai)

        else:
            # Paket 2, 3, 4: butuh deteksi zona → Phase 1-3 dulu
            roman_start_p, bab_p_list = proc.scan_zones()
            detected_bab_texts = [
                DocProcessor._p_text(p)[:60] for p in bab_p_list
            ]
            # Geser roman_start_p jika user memiliki lebih dari 1 cover
            if num_cover > 1 and hidden_cov == 'Ya':
                roman_start_p = DocProcessor.advance_roman_start(
                    doc, roman_start_p, num_cover
                )
            roman_start_p             = proc.insert_breaks(roman_start_p, bab_p_list)
            roman_sec, bab_sec_list, n_sections = proc.build_section_map(
                roman_start_p, bab_p_list
            )

            if paket == 'paket2':
                paket2.apply(proc, roman_sec, bab_sec_list, n_sections, hidden_cov, posisi,
                             dimulai_dari=dimulai)
            elif paket == 'paket4':
                paket4.apply(
                    proc, roman_sec, bab_sec_list, n_sections, hidden_cov,
                    pos_romawi=posisi, pos_bab=pos_bab, pos_isi_bab=pos_isi_bab,
                    dimulai_dari=dimulai, semb_dafus=semb_dafus, semb_lamprn=semb_lamprn,
                    bab_p_list=bab_p_list
                )
            else:
                paket3.apply(proc, roman_sec, bab_sec_list, n_sections, hidden_cov,
                             dimulai_dari=dimulai)

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
