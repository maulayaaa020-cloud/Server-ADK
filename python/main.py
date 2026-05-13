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

# Pastikan folder python/ ada di sys.path agar import utils/paket* bisa berjalan
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import paket1
import paket2
import paket3
from utils import DocProcessor


def _fail(code, message):
    print(json.dumps({"status": "error", "code": code, "message": message}))
    sys.exit(1)


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

        if paket == 'paket1':
            # Paket 1: tidak butuh deteksi zona
            paket1.apply(proc, hidden_cov, posisi)

        else:
            # Paket 2 & 3: butuh deteksi zona → Phase 1-3 dulu
            roman_start_p, bab_p_list = proc.scan_zones()
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
        "sections":       sections_info
    }))


if __name__ == '__main__':
    main()
