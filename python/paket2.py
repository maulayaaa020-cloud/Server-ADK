"""
paket2.py — Romawi + Angka, posisi bebas seragam.
Semua zona (cover, romawi, BAB) menggunakan posisi yang sama (dipilih user).
"""


def apply(proc, roman_sec, bab_sec_list, n_sections, hidden_cov, posisi):
    """
    proc        : DocProcessor instance
    roman_sec   : index section awal zona romawi (None jika tidak ada)
    bab_sec_list: list index section tiap heading BAB
    n_sections  : total section dalam doc
    hidden_cov  : 'Ya' | 'Tidak'
    posisi      : str, e.g. 'Tengah Bawah'
    """
    align         = proc._get_h_align(posisi)
    top           = proc._is_top(posisi)
    cov_show      = None if hidden_cov == 'Ya' else (align, top)
    first_bab_sec = bab_sec_list[0] if bab_sec_list else None

    for i, section in enumerate(proc.doc.sections):

        # ── Cover zone ──
        if roman_sec is not None and i < roman_sec:
            proc.fmt_cover(section, first_cover=(i == 0), show_pos=cov_show,
                           visible_pos=(align, top))
            continue
        if roman_sec is None and i == 0:
            proc.fmt_cover(section, first_cover=True, show_pos=cov_show,
                           visible_pos=(align, top))
            continue

        # ── Romawi zone ──
        if first_bab_sec is None or i < first_bab_sec:
            proc.fmt_uniform_section(section, align, top, fmt='lowerRoman')
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

        proc.fmt_uniform_section(
            section, align, top, fmt='decimal',
            start=1 if (is_zone_first and zone_idx == 0) else None
        )
