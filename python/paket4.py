"""
paket4.py — Romawi + Angka, full custom positions per zona.
  Cover    : tanpa nomor (atau dengan nomor jika hidden_cov='Tidak')
  Romawi   : posisi + letak dipilih user (pos_romawi)
  BAB first: posisi + letak dipilih user (pos_bab)
  BAB cont : posisi + letak dipilih user (pos_isi_bab)
  Dafpus   : tanpa nomor jika semb_dafus='Ya'
  Lampiran : tanpa nomor jika semb_lamprn='Ya'
  Dimulai  : nomor awal romawi (i/ii/iii/iv)
"""
import re
from docx.oxml.ns import qn
from docx.shared import Cm


def _get_text(p_elem):
    return ''.join(t.text or '' for t in p_elem.findall('.//' + qn('w:t'))).strip()


def _is_dafpus_text(text):
    return bool(re.match(
        r'^\s*(daftar\s+pust?aka|referensi|references?|bibliography|'
        r'bibliographies|works?\s+cited|literature\s+cited)\s*$',
        text, re.IGNORECASE
    ))


def _is_lampiran_text(text):
    return bool(re.match(
        r'^\s*(lampiran|appendix|appendices|attachment)',
        text, re.IGNORECASE
    ))


_DIVIDER_WORDS = ('LAMPIRAN', 'APPENDIX', 'APPENDICES', 'ATTACHMENT')


def _find_lampiran_letter_divider(doc, first_lsec):
    """Temukan elemen paragraf pertama dari halaman pemisah LAMPIRAN/APPENDIX
    yang berada tepat sebelum seksi lampiran pertama (first_lsec).

    Mendukung dua pola:
    1. Huruf-per-baris  : L / A / M / P / I / R / A / N  (atau A/P/P/E/N/D/I/X)
    2. Satu kata penuh  : paragraf tunggal berisi "LAMPIRAN" atau "APPENDIX"
       tanpa konten lain, yang berada tepat sebelum batas seksi.

    Kembalikan elemen XML paragraf pertama dari divider, atau None.
    """
    W    = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    body = list(doc.element.body)

    # Temukan paragraf ke-(first_lsec)-th yang punya sectPr — batas seksi.
    sect_count   = 0
    boundary_idx = -1
    for j, el in enumerate(body):
        if not el.tag.endswith('}p'):
            continue
        pPr = el.find('{%s}pPr' % W)
        if pPr is not None and pPr.find('{%s}sectPr' % W) is not None:
            sect_count += 1
            if sect_count == first_lsec:
                boundary_idx = j
                break
    if boundary_idx < 0:
        return None

    def _txt(el):
        return ''.join(t.text or '' for t in el.iter('{%s}t' % W)).strip()

    # ── Pola 1: huruf-per-baris ─────────────────────────────────────────
    letters = []
    for j in range(boundary_idx, max(boundary_idx - 12, -1), -1):
        el = body[j]
        if not el.tag.endswith('}p'):
            break
        t = _txt(el)
        if len(t) == 1 and t.isalpha():
            letters.insert(0, (j, el))
        else:
            break
    if letters:
        concat = ''.join(_txt(el) for _, el in letters).upper()
        if any(w in concat or concat in w for w in _DIVIDER_WORDS):
            return letters[0][1]

    # ── Pola 2: satu kata penuh (misal "APPENDIX" pada halaman tersendiri) ─
    # Periksa 3 paragraf terakhir seksi sebelum batas: jika ada yang teksnya
    # persis salah satu kata divider (tanpa imbuhan nomor), kembalikan itu.
    for j in range(boundary_idx, max(boundary_idx - 4, -1), -1):
        el = body[j]
        if not el.tag.endswith('}p'):
            continue
        t = _txt(el).upper()
        if t in _DIVIDER_WORDS:
            # Pastikan ini bukan paragraf dengan konten lain (nomor halaman dll)
            # dan berada di akhir seksi (tidak jauh dari batas)
            return el

    return None


def _blank_section(proc, section):
    section.different_first_page_header_footer = False
    section.header_distance = Cm(1.25)
    section.footer_distance = Cm(1.25)
    proc.clear_header(section)
    proc.clear_footer(section)


def apply(proc, roman_sec, bab_sec_list, n_sections, hidden_cov,
          pos_romawi='Tengah Bawah', pos_bab='Tengah Bawah',
          pos_isi_bab='Kanan Atas', dimulai_dari='i',
          semb_dafus='Tidak', semb_lamprn='Tidak',
          bab_p_list=None):
    """
    proc         : DocProcessor instance
    roman_sec    : index section awal zona romawi (None jika tidak ada)
    bab_sec_list : list index section tiap heading BAB/Dafpus/Lampiran
    n_sections   : total section dalam doc
    hidden_cov   : 'Ya' | 'Tidak'
    pos_romawi   : posisi halaman romawi, e.g. 'Tengah Bawah'
    pos_bab      : posisi halaman pertama BAB, e.g. 'Tengah Bawah'
    pos_isi_bab  : posisi halaman lanjutan BAB, e.g. 'Kanan Atas'
    dimulai_dari : 'i' | 'ii' | 'iii' | 'iv'
    semb_dafus   : 'Ya' | 'Tidak'
    semb_lamprn  : 'Ya' | 'Tidak'
    bab_p_list   : list paragraph elements dari scan_zones (untuk deteksi dafpus/lampiran)
    """
    roman_start_map = {'i': 1, 'ii': 2, 'iii': 3, 'iv': 4}
    roman_start_num = roman_start_map.get((dimulai_dari or 'i').lower().strip(), 1)

    align_romawi = proc._get_h_align(pos_romawi)
    top_romawi   = proc._is_top(pos_romawi)
    align_bab    = proc._get_h_align(pos_bab)
    top_bab      = proc._is_top(pos_bab)
    align_isi    = proc._get_h_align(pos_isi_bab)
    top_isi      = proc._is_top(pos_isi_bab)

    cov_show      = None if hidden_cov == 'Ya' else (align_romawi, top_romawi)
    first_bab_sec = bab_sec_list[0] if bab_sec_list else None
    first_roman_done = False

    # Identifikasi section yang termasuk dafpus / lampiran
    dafpus_secs  = set()
    lampiran_secs = set()
    if bab_p_list:
        for k, p_elem in enumerate(bab_p_list):
            if k >= len(bab_sec_list):
                break
            text      = _get_text(p_elem)
            sec_start = bab_sec_list[k]
            sec_end   = bab_sec_list[k + 1] if k + 1 < len(bab_sec_list) else n_sections
            if _is_dafpus_text(text):
                for si in range(sec_start, sec_end):
                    dafpus_secs.add(si)
            elif _is_lampiran_text(text):
                for si in range(sec_start, sec_end):
                    lampiran_secs.add(si)

    # Jika semb_lamprn='Ya': cek apakah ada halaman pemisah LAMPIRAN bergaya huruf-per-baris
    # (L, A, M, P, I, R, A, N) tepat sebelum seksi lampiran pertama.
    # Jika ada, sisipkan section break sebelum huruf pertama agar halaman pemisah
    # mendapat seksinya sendiri dan bisa disembunyikan (masuk lampiran_secs).
    if semb_lamprn == 'Ya' and lampiran_secs:
        _first_lsec   = min(lampiran_secs)
        _divider_elem = _find_lampiran_letter_divider(proc.doc, _first_lsec)
        if _divider_elem is not None:
            proc.insert_break_before_xml(_divider_elem)
            # Geser semua referensi seksi >= _first_lsec ke atas 1
            bab_sec_list  = [(s + 1 if s >= _first_lsec else s) for s in bab_sec_list]
            lampiran_secs = {(s + 1 if s >= _first_lsec else s) for s in lampiran_secs}
            lampiran_secs.add(_first_lsec)        # tambahkan seksi pemisah baru
            dafpus_secs   = {(s + 1 if s >= _first_lsec else s) for s in dafpus_secs}
            n_sections   += 1
            first_bab_sec = bab_sec_list[0] if bab_sec_list else first_bab_sec

    # cover_start: nilai page pertama cover section.
    # hidden_cov='Ya'  → cover 1 tersembunyi, cover 2+ mulai dari roman_start_num
    #                    sehingga cover section dimulai dari roman_start_num - 1
    # hidden_cov='Tidak' → cover 1 ditampilkan, mulai dari roman_start_num
    cover_start = roman_start_num - 1 if hidden_cov == 'Ya' else roman_start_num

    for i, section in enumerate(proc.doc.sections):
        try:
            # ── Cover zone ──────────────────────────────────────────────────
            if roman_sec is not None and i < roman_sec:
                proc.fmt_cover(section, first_cover=(i == 0), show_pos=cov_show,
                               cover_start=cover_start,
                               visible_pos=(align_romawi, top_romawi))
                continue
            if roman_sec is None and i == 0:
                proc.fmt_cover(section, first_cover=True, show_pos=cov_show,
                               cover_start=cover_start,
                               visible_pos=(align_romawi, top_romawi))
                continue

            # ── Romawi zone ─────────────────────────────────────────────────
            if first_bab_sec is None or i < first_bab_sec:
                if not first_roman_done and hidden_cov == 'Ya':
                    proc.fmt_uniform_section(
                        section, align_romawi, top_romawi,
                        fmt='lowerRoman', start=roman_start_num
                    )
                    first_roman_done = True
                else:
                    proc.fmt_uniform_section(
                        section, align_romawi, top_romawi,
                        fmt='lowerRoman', start=None
                    )
                continue

            # ── BAB zone ────────────────────────────────────────────────────
            zone_idx      = -1
            is_zone_first = False
            for k, bab_sec in enumerate(bab_sec_list):
                next_sec = bab_sec_list[k + 1] if k + 1 < len(bab_sec_list) else n_sections
                if bab_sec <= i < next_sec:
                    zone_idx      = k
                    is_zone_first = (i == bab_sec)
                    break

            # Sembunyikan nomor pada dafpus / lampiran jika dipilih
            if semb_dafus == 'Ya' and i in dafpus_secs:
                _blank_section(proc, section)
                continue
            if semb_lamprn == 'Ya' and i in lampiran_secs:
                _blank_section(proc, section)
                continue

            # Halaman pertama BAB dengan posisi berbeda dari lanjutan
            if is_zone_first and pos_bab != pos_isi_bab:
                section.different_first_page_header_footer = True
                section.header_distance = Cm(1.25)
                section.footer_distance = Cm(1.25)
                proc.clear_header(section)
                proc.clear_footer(section)

                fph = section.first_page_header
                fph.is_linked_to_previous = False
                fpf = section.first_page_footer
                fpf.is_linked_to_previous = False
                for p in list(fph.paragraphs):
                    proc.clear_paragraph(p)
                for p in list(fpf.paragraphs):
                    proc.clear_paragraph(p)

                if top_bab:
                    p = proc._first_para(fph)
                    proc.clear_paragraph(p)
                    p.alignment = align_bab
                    proc.add_page_number(p)
                    proc._set_pn_spacing(p)
                else:
                    p = proc._first_para(fpf)
                    proc.clear_paragraph(p)
                    p.alignment = align_bab
                    proc.add_page_number(p)
                    proc._set_pn_spacing(p)

                if top_isi:
                    proc._place_num_in_part(section.header, align_isi)
                else:
                    proc._place_num_in_part(section.footer, align_isi)

                proc.set_page_number_format(
                    section, 'decimal', 1 if zone_idx == 0 else None
                )

            elif is_zone_first:
                # Posisi pertama dan lanjutan sama
                proc.fmt_uniform_section(
                    section, align_bab, top_bab,
                    fmt='decimal', start=1 if zone_idx == 0 else None
                )

            else:
                # Halaman lanjutan BAB
                proc.fmt_uniform_section(section, align_isi, top_isi, fmt='decimal')

        except Exception:
            pass
