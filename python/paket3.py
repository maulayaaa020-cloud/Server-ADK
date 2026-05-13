"""
paket3.py — Romawi + Angka, posisi tetap (Populer Skripsi).
  Cover    : tanpa nomor (atau romawi jika hidden_cov='Tidak')
  Romawi   : tengah bawah (lowerRoman)
  BAB first: bawah tengah (halaman 1 BAB) + kanan atas (halaman selanjutnya)
  BAB cont : kanan atas (semua halaman, tanpa different_first_page)
"""
from docx.enum.text import WD_ALIGN_PARAGRAPH


def apply(proc, roman_sec, bab_sec_list, n_sections, hidden_cov):
    """
    proc        : DocProcessor instance
    roman_sec   : index section awal zona romawi
    bab_sec_list: list index section tiap heading BAB
    n_sections  : total section
    hidden_cov  : 'Ya' | 'Tidak'
    """
    cov_show      = None if hidden_cov == 'Ya' else (WD_ALIGN_PARAGRAPH.CENTER, False)
    first_bab_sec = bab_sec_list[0] if bab_sec_list else None

    for i, section in enumerate(proc.doc.sections):

        # ── Cover zone ──
        if roman_sec is not None and i < roman_sec:
            proc.fmt_cover(section, first_cover=(i == 0), show_pos=cov_show)
            continue
        if roman_sec is None and i == 0:
            proc.fmt_cover(section, first_cover=True, show_pos=cov_show)
            continue

        # ── Romawi zone ──
        if first_bab_sec is None or i < first_bab_sec:
            proc.fmt_roman(section)
            continue

        # ── BAB zone ──
        zone_idx      = -1
        is_zone_first = False
        for k, bab_sec in enumerate(bab_sec_list):
            next_sec = bab_sec_list[k + 1] if k + 1 < len(bab_sec_list) else n_sections
            if bab_sec <= i < next_sec:
                zone_idx      = k
                is_zone_first = (i == bab_sec)
                break

        if is_zone_first:
            proc.fmt_bab_first(section, reset_to_1=(zone_idx == 0))
        else:
            proc.fmt_bab_continuation(section)
