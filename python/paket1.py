"""
paket1.py — Full Angka (posisi bebas, semua halaman bernomor Arab).
Tidak butuh deteksi zona; langsung terapkan ke semua section.
"""
from docx.shared import Cm


def apply(proc, hidden_cov, posisi):
    """
    proc      : DocProcessor instance
    hidden_cov: 'Ya' (cover tanpa nomor) | 'Tidak' (cover tampil nomor)
    posisi    : str, e.g. 'Tengah Bawah', 'Kanan Atas', dll.
    """
    align = proc._get_h_align(posisi)
    top   = proc._is_top(posisi)

    for i, section in enumerate(proc.doc.sections):
        section.header_distance = Cm(1.25)
        section.footer_distance = Cm(1.25)

        if i == 0 and hidden_cov == 'Ya':
            # Cover: beda halaman pertama → kosongkan header/footer cover
            section.different_first_page_header_footer = True
            for part in [section.first_page_header, section.first_page_footer]:
                part.is_linked_to_previous = False
                for p in part.paragraphs:
                    proc.clear_paragraph(p)
            proc.set_page_number_format(section, 'decimal', 1)
        else:
            section.different_first_page_header_footer = False
            proc.set_page_number_format(section, 'decimal', 1 if i == 0 else None)

        if top:
            proc.clear_footer(section)
            proc._place_num_in_part(section.header, align)
        else:
            proc.clear_header(section)
            proc._place_num_in_part(section.footer, align)
