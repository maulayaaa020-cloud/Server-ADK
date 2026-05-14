"""
utils.py — Shared helpers untuk semua paket penomoran halaman.
Semua fungsi XML/docx, deteksi zona, dan section formatting ada di sini.
"""
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Pt, Cm
from copy import deepcopy
import re


# =========================================================
# DETECTION — module-level (tidak butuh doc/font)
# =========================================================

ROMAN_START_KEYWORDS = [
    "kata pengantar", "prakata", "ucapan terima kasih",
    "abstrak",
    "lembar pengesahan", "lembar persetujuan", "halaman pengesahan",
    "pengesahan", "persetujuan",
    "lembar pernyataan", "pernyataan keaslian",
    "halaman pernyataan",
    "rekomendasi",
    "halaman persembahan", "persembahan", "motto",
    "ringkasan",
    "daftar isi",
    "abstract", "summary", "executive summary",
    "preface", "foreword", "acknowledgment", "acknowledgements",
    "approval page", "approval sheet", "declaration", "originality statement",
    "dedication",
    "table of contents", "list of contents", "contents",
    "list of tables", "list of figures", "list of appendices",
]

_ROMAN_PAT = r'm{0,4}(cm|cd|d?c{0,3})(xc|xl|l?x{0,3})(ix|iv|v?i{0,3})'
BAB_HEAD_RE = re.compile(
    rf'^\s*(bab|chapter)\s+({_ROMAN_PAT}|\d+)\b(.*?)$',
    re.IGNORECASE | re.DOTALL
)


def is_roman_start(text):
    lower = text.strip().lower()
    return any(lower == k or lower.startswith(k) for k in ROMAN_START_KEYWORDS)


def is_bab_heading(text):
    text = text.strip()
    if not text:
        return False
    if BAB_HEAD_RE.match(text):
        return True
    if re.match(
        r'^\s*(daftar\s+pust?aka|referensi|references?|reference\s+list|bibliography|'
        r'bibliographies|works?\s+cited|literature\s+cited)\s*$', text, re.IGNORECASE
    ):
        return True
    if re.match(r'^\s*(lampiran|appendix|appendices|attachment)(\s+.*)?$', text, re.IGNORECASE):
        return True
    return False


def is_false_bab(para):
    text  = para.text.strip()
    style = para.style.name.lower() if para.style else ""
    m = BAB_HEAD_RE.match(text)
    if m:
        sisa_raw = m.group(6) or ""
        sisa = sisa_raw.strip()
        if not sisa_raw.startswith('\n'):
            if len(sisa) > 60:
                return True
            if len(text.split()) > 8:
                return True
        non_heading = ('body', 'list', 'quote', 'caption', 'web')
        if sisa and any(s in style for s in non_heading):
            return True
    for pattern in [
        r'^\s*(daftar\s+pust?aka|referensi|references?|reference\s+list|bibliography|'
        r'bibliographies|works?\s+cited|literature\s+cited)\s+\S',
        r'^\s*(lampiran|appendix|appendices|attachment)\s+[^\n]{60,}',
    ]:
        if re.match(pattern, text, re.IGNORECASE):
            return True
    if re.match(r'^\s*(lampiran|appendix|appendices|attachment)', text, re.IGNORECASE):
        if (re.search(r'\t\s*\d+\s*$', text) or
                re.search(r'\.{3,}.*\d+\s*$', text) or
                re.search(r'\s{4,}\d+\s*$', text)):
            return True
    # LAMPIRAN / DAFTAR PUSTAKA sebagai sub-heading (Heading 2+) bukan section mandiri
    _endpoint_pat = (
        r'^\s*(lampiran|appendix|appendices|attachment|'
        r'daftar\s+pust?aka|referensi|references?|bibliography)'
    )
    if re.match(_endpoint_pat, text, re.IGNORECASE):
        if re.search(r'heading\s*[2-9]', style, re.IGNORECASE):
            return True
    return False


def is_toc_heading(text):
    lower = text.strip().lower()
    return any(k in lower for k in [
        "daftar isi", "daftar tabel", "daftar gambar", "daftar lampiran",
        "table of contents", "list of contents", "list of tables",
        "list of figures", "list of appendices",
    ])


def is_toc_entry(text):
    # Multi-line (soft return): TOC entry hanya jika ada pola nomor halaman di salah satu baris.
    # "BAB I\nPENDAHULUAN" tidak punya nomor → bukan TOC entry.
    if '\n' in text:
        for line in text.split('\n'):
            if re.search(r'\t\s*[\divxlcdmIVXLCDM]+\s*$', line):
                return True
            if re.search(r'\.{3,}.*\d+\s*$', line):
                return True
        return False
    if re.search(r'\t\s*[\divxlcdmIVXLCDM]+\s*$', text):
        return True
    if re.search(r'\.{3,}.*\d+\s*$', text):
        return True
    if re.search(r'\s{4,}\d+\s*$', text):
        return True
    return False


# =========================================================
# DocProcessor — menyimpan doc + font settings, semua metode
# =========================================================

class DocProcessor:

    def __init__(self, doc, font_name, font_size):
        self.doc       = doc
        self.font_name = font_name
        self.font_size = font_size  # int (pt)

    # ── Low-level helpers ─────────────────────────────────

    def add_page_number(self, paragraph):
        run = paragraph.add_run()
        run.font.name = self.font_name
        run.font.size = Pt(self.font_size)
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')
        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')
        instrText.text = " PAGE "
        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')
        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)

    @staticmethod
    def _first_para(part):
        if not part.paragraphs:
            part._element.append(OxmlElement('w:p'))
        return part.paragraphs[0]

    @staticmethod
    def clear_paragraph(paragraph):
        p = paragraph._p
        for child in list(p):
            p.remove(child)

    @staticmethod
    def _set_pn_spacing(p):
        pf = p.paragraph_format
        pf.space_after        = Pt(0)
        pf.space_before       = Pt(0)
        pf.line_spacing_rule  = WD_LINE_SPACING.SINGLE

    @staticmethod
    def _get_h_align(pos):
        if 'Tengah' in pos: return WD_ALIGN_PARAGRAPH.CENTER
        if 'Kanan'  in pos: return WD_ALIGN_PARAGRAPH.RIGHT
        return WD_ALIGN_PARAGRAPH.LEFT

    @staticmethod
    def _is_top(pos):
        return 'Atas' in pos

    def _place_num_in_part(self, part, align):
        part.is_linked_to_previous = False
        for p in part.paragraphs:
            self.clear_paragraph(p)
        p = self._first_para(part)
        self.clear_paragraph(p)
        p.alignment = align
        self.add_page_number(p)
        self._set_pn_spacing(p)

    # ── Header / Footer helpers ───────────────────────────

    def purge_all_headers_footers(self):
        for section in self.doc.sections:
            for part in [
                section.header, section.footer,
                section.first_page_header, section.first_page_footer,
            ]:
                part.is_linked_to_previous = False
                elem = part._element
                for sdt in list(elem.findall(qn('w:sdt'))):
                    elem.remove(sdt)
                paras = list(elem.findall(qn('w:p')))
                if paras:
                    for child in list(paras[0]):
                        paras[0].remove(child)
                    for extra in paras[1:]:
                        elem.remove(extra)
                else:
                    elem.append(OxmlElement('w:p'))

    @staticmethod
    def clear_header(section):
        h = section.header
        h.is_linked_to_previous = False
        for p in h.paragraphs:
            DocProcessor.clear_paragraph(p)

    @staticmethod
    def clear_footer(section):
        f = section.footer
        f.is_linked_to_previous = False
        for p in f.paragraphs:
            DocProcessor.clear_paragraph(p)

    @staticmethod
    def set_page_number_format(section, fmt='decimal', start=None):
        sectPr = section._sectPr
        for old in sectPr.findall(qn('w:pgNumType')):
            sectPr.remove(old)
        pgNumType = OxmlElement('w:pgNumType')
        pgNumType.set(qn('w:fmt'), fmt)
        if start is not None:
            pgNumType.set(qn('w:start'), str(start))
        sectPr.append(pgNumType)

    # ── XML section-break helpers ─────────────────────────

    @staticmethod
    def _p_text(p_elem):
        return ''.join(t.text or '' for t in p_elem.findall('.//' + qn('w:t'))).strip()

    @staticmethod
    def _p_has_content(p_elem):
        if DocProcessor._p_text(p_elem):
            return True
        if p_elem.find('.//' + qn('w:drawing')) is not None:
            return True
        if p_elem.find('.//' + qn('w:pict')) is not None:
            return True
        if p_elem.find('.//' + qn('w:fldChar')) is not None:
            return True
        return False

    @staticmethod
    def _has_sectPr(p_elem):
        pPr = p_elem.find(qn('w:pPr'))
        return pPr is not None and pPr.find(qn('w:sectPr')) is not None

    def _make_sectPr(self):
        new_sectPr = deepcopy(self.doc.sections[-1]._sectPr)
        for child in list(new_sectPr):
            tag = child.tag
            if tag.endswith('headerReference') or tag.endswith('footerReference'):
                new_sectPr.remove(child)
        type_elem = new_sectPr.find(qn('w:type'))
        if type_elem is None:
            type_elem = OxmlElement('w:type')
            new_sectPr.append(type_elem)
        type_elem.set(qn('w:val'), 'nextPage')
        return new_sectPr

    def _attach_sectPr(self, p_elem):
        pPr = p_elem.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr')
            p_elem.insert(0, pPr)
        pPr.append(self._make_sectPr())

    def _purge_continuous_sectPr_in_bab_zone(self, first_bab_p):
        body     = self.doc.element.body
        children = list(body)
        try:
            start_idx = children.index(first_bab_p)
        except ValueError:
            return
        for elem in children[start_idx:]:
            if not elem.tag.endswith('}p'):
                continue
            pPr = elem.find(qn('w:pPr'))
            if pPr is None:
                continue
            sectPr = pPr.find(qn('w:sectPr'))
            if sectPr is None:
                continue
            type_e = sectPr.find(qn('w:type'))
            if type_e is None or type_e.get(qn('w:val')) != 'continuous':
                continue
            pgsz = sectPr.find(qn('w:pgSz'))
            if pgsz is not None and pgsz.get(qn('w:orient')) == 'landscape':
                continue
            pPr.remove(sectPr)

    def _strip_empty_paras_before_bab(self, target_p):
        body     = self.doc.element.body
        children = list(body)
        try:
            tgt_idx = children.index(target_p)
        except ValueError:
            return
        to_remove = []
        for j in range(tgt_idx - 1, -1, -1):
            elem = children[j]
            if not elem.tag.endswith('}p'):
                break
            pPr = elem.find(qn('w:pPr'))
            if pPr is not None and pPr.find(qn('w:sectPr')) is not None:
                break
            if self._p_has_content(elem):
                break
            to_remove.append(elem)
        for elem in to_remove:
            body.remove(elem)

    def _remove_page_breaks_before(self, target_p):
        body     = self.doc.element.body
        children = list(body)
        idx      = children.index(target_p)
        for br in target_p.findall('.//' + qn('w:br')):
            if br.get(qn('w:type')) == 'page':
                parent = br.getparent()
                if parent is not None:
                    parent.remove(br)
        for j in range(idx - 1, -1, -1):
            elem = children[j]
            if not elem.tag.endswith('}p'):
                break
            if self._p_has_content(elem):
                break
            for br in elem.findall('.//' + qn('w:br')):
                if br.get(qn('w:type')) == 'page':
                    parent = br.getparent()
                    if parent is not None:
                        parent.remove(br)

    def insert_break_before_xml(self, target_p):
        self._remove_page_breaks_before(target_p)
        body     = self.doc.element.body
        children = list(body)
        tgt_idx  = children.index(target_p)
        j = tgt_idx - 1
        while j >= 0:
            elem = children[j]
            tag  = elem.tag
            if tag.endswith('}p'):
                pPr = elem.find(qn('w:pPr'))
                if pPr is not None:
                    sectPr = pPr.find(qn('w:sectPr'))
                    if sectPr is not None:
                        if sectPr.find(qn('w:pgSz')) is not None:
                            return
                        else:
                            pPr.remove(sectPr)
                            if self._p_has_content(elem):
                                self._attach_sectPr(elem)
                                return
                            j -= 1
                            continue
                if self._p_has_content(elem):
                    self._attach_sectPr(elem)
                    return
            elif tag.endswith('}tbl') or tag.endswith('}sdt'):
                new_p   = OxmlElement('w:p')
                new_pPr = OxmlElement('w:pPr')
                new_p.append(new_pPr)
                new_pPr.append(self._make_sectPr())
                body.insert(tgt_idx, new_p)
                return
            j -= 1

    def _find_section_start(self, target_p):
        body     = self.doc.element.body
        children = list(body)
        try:
            idx = children.index(target_p)
        except ValueError:
            return target_p
        for j in range(idx - 1, -1, -1):
            elem = children[j]
            if elem.tag.endswith('}p') and self._has_sectPr(elem):
                for k in range(j + 1, idx + 1):
                    c = children[k]
                    if c.tag.endswith('}p'):
                        return c
                break
        return target_p

    @staticmethod
    def _has_page_break_before(paras, from_idx, to_idx):
        for para in paras[from_idx:to_idx]:
            if DocProcessor._has_sectPr(para._p):
                return True
            for br in para._p.findall('.//' + qn('w:br')):
                if br.get(qn('w:type')) == 'page':
                    return True
        return False

    @staticmethod
    def _para_has_page_break_before(para):
        pPr = para._p.find(qn('w:pPr'))
        if pPr is None:
            return False
        pb = pPr.find(qn('w:pageBreakBefore'))
        if pb is None:
            return False
        return pb.get(qn('w:val'), 'true') not in ('false', '0')

    # ── Phase pipeline ────────────────────────────────────

    def scan_zones(self):
        """Phase 1: Scan paragraphs. Returns (roman_start_p, bab_p_list)."""
        roman_start_p      = None
        bab_p_list         = []
        lampiran_found     = False
        last_bab_para_idx  = None
        found_numbered_bab = False
        seen_bab_numbers   = set()   # cegah BAB dengan nomor sama terdeteksi dua kali
        inside_toc         = False
        all_paras          = list(self.doc.paragraphs)

        for para_idx, para in enumerate(all_paras):
            text  = para.text.strip()
            lower = text.lower()

            if is_toc_heading(lower):
                inside_toc = True
                if roman_start_p is None:
                    roman_start_p = para._p
                continue

            if inside_toc:
                if is_toc_entry(text):
                    continue
                else:
                    inside_toc = False

            if not text:
                continue

            if roman_start_p is None and is_roman_start(text):
                roman_start_p = para._p

            if is_bab_heading(text) and not is_false_bab(para):
                is_numbered_bab = BAB_HEAD_RE.match(text) is not None
                if not is_numbered_bab and not found_numbered_bab:
                    continue
                if re.match(r'^\s*lampiran', text, re.IGNORECASE):
                    if lampiran_found:
                        continue
                    lampiran_found = True
                if is_numbered_bab:
                    found_numbered_bab = True
                    m_num = BAB_HEAD_RE.match(text)
                    bab_num_key = m_num.group(2).strip().lower() if m_num else None
                    if bab_num_key and bab_num_key in seen_bab_numbers:
                        continue  # Nomor BAB sudah terdeteksi → duplikat, skip
                else:
                    bab_num_key = None
                if last_bab_para_idx is not None:
                    _window = max(last_bab_para_idx + 1, para_idx - 15)
                    has_break = self._has_page_break_before(all_paras, _window, para_idx)
                    if not has_break:
                        has_break = self._para_has_page_break_before(para)
                    # "Heading X" style + page break → langsung dipercaya.
                    # Non-heading style + page break → tetap cek forward_count karena sectPr
                    # di body text (misal Sistematika Penulisan) bisa memicu false positive.
                    _style_lower = (para.style.name.lower() if para.style else "")
                    _trusted = has_break and bool(re.search(r'heading', _style_lower))
                    if not _trusted:
                        # Lampiran & Daftar Pustaka dibebaskan dari forward_count
                        _is_endpoint = bool(re.match(
                            r'^\s*(lampiran|daftar\s+pust?aka)', text, re.IGNORECASE
                        ))
                        if not _is_endpoint:
                            forward_count = 0
                            for _k in range(para_idx + 1, len(all_paras)):
                                _nt = all_paras[_k].text.strip()
                                if is_bab_heading(_nt) and not is_false_bab(all_paras[_k]):
                                    break
                                if _nt:
                                    forward_count += 1
                            # Threshold lebih rendah (2) jika ada page break, karena
                            # Sistematika entries hanya punya 1 kalimat per BAB.
                            _threshold = 2 if has_break else 5
                            if forward_count < _threshold:
                                continue
                bab_p_list.append(para._p)
                if bab_num_key:
                    seen_bab_numbers.add(bab_num_key)
                last_bab_para_idx = para_idx

        return roman_start_p, bab_p_list

    def insert_breaks(self, roman_start_p, bab_p_list):
        """Phase 1.5 + 2: Purge continuous sectPr, insert zone breaks.
        Returns updated roman_start_p (may differ after _find_section_start)."""
        if bab_p_list:
            self._purge_continuous_sectPr_in_bab_zone(bab_p_list[0])

        if roman_start_p is not None:
            roman_start_p = self._find_section_start(roman_start_p)
            self.insert_break_before_xml(roman_start_p)

        for bab_p in bab_p_list:
            self.insert_break_before_xml(bab_p)
            self._strip_empty_paras_before_bab(bab_p)

        return roman_start_p  # dikembalikan karena bisa berubah

    def build_section_map(self, roman_start_p, bab_p_list):
        """Phase 3: Build section boundary map.
        Returns (roman_sec, bab_sec_list, n_sections)."""
        breaks = []
        for para in self.doc.paragraphs:
            if self._has_sectPr(para._p):
                breaks.append(para._p)

        para_list   = list(self.doc.paragraphs)
        para_p_list = [p._p for p in para_list]
        break_indices = [para_p_list.index(bp) for bp in breaks]
        boundaries    = [0] + [bi + 1 for bi in break_indices]
        n_sections    = len(boundaries)

        def para_to_sec(p_elem):
            try:
                idx = para_p_list.index(p_elem)
            except ValueError:
                return n_sections - 1
            for s in range(n_sections):
                start = boundaries[s]
                end   = boundaries[s + 1] if s + 1 < n_sections else len(para_p_list)
                if start <= idx < end:
                    return s
            return n_sections - 1

        roman_sec    = para_to_sec(roman_start_p) if roman_start_p is not None else None
        bab_sec_list = [para_to_sec(p) for p in bab_p_list]
        return roman_sec, bab_sec_list, n_sections

    # ── Section formatters (dipakai oleh paket2 & paket3) ─

    def fmt_cover(self, section, first_cover=False, show_pos=None):
        self.clear_header(section)
        self.clear_footer(section)
        if first_cover:
            self.set_page_number_format(section, 'lowerRoman', 1)
        else:
            self.set_page_number_format(section, 'lowerRoman')
        if show_pos and first_cover:
            align, top = show_pos
            if top:
                self._place_num_in_part(section.header, align)
            else:
                self._place_num_in_part(section.footer, align)

    def fmt_roman(self, section):
        section.different_first_page_header_footer = False
        section.footer_distance = Cm(1.25)
        self.clear_header(section)
        f = section.footer
        f.is_linked_to_previous = False
        p = self._first_para(f)
        self.clear_paragraph(p)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self.add_page_number(p)
        self._set_pn_spacing(p)
        self.set_page_number_format(section, 'lowerRoman')

    def fmt_bab_first(self, section, reset_to_1=False):
        section.different_first_page_header_footer = True
        section.header_distance = Cm(1.25)
        section.footer_distance = Cm(1.25)
        self.clear_header(section)
        self.clear_footer(section)

        fph = section.first_page_header
        fph.is_linked_to_previous = False
        for p in fph.paragraphs:
            self.clear_paragraph(p)

        fpf = section.first_page_footer
        fpf.is_linked_to_previous = False
        for p in fpf.paragraphs:
            self.clear_paragraph(p)
        p = self._first_para(fpf)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self.add_page_number(p)
        self._set_pn_spacing(p)

        h = section.header
        h.is_linked_to_previous = False
        p = self._first_para(h)
        self.clear_paragraph(p)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        self.add_page_number(p)
        self._set_pn_spacing(p)

        if reset_to_1:
            self.set_page_number_format(section, 'decimal', 1)
        else:
            self.set_page_number_format(section, 'decimal')

    def fmt_bab_continuation(self, section):
        section.different_first_page_header_footer = False
        section.header_distance = Cm(1.25)
        section.footer_distance = Cm(1.25)
        self.clear_header(section)
        self.clear_footer(section)
        h = section.header
        h.is_linked_to_previous = False
        p = self._first_para(h)
        self.clear_paragraph(p)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        self.add_page_number(p)
        self._set_pn_spacing(p)
        self.set_page_number_format(section, 'decimal')

    def fmt_uniform_section(self, section, align, top, fmt='decimal', start=None):
        section.header_distance = Cm(1.25)
        section.footer_distance = Cm(1.25)
        section.different_first_page_header_footer = False
        self.clear_header(section)
        self.clear_footer(section)
        if top:
            self._place_num_in_part(section.header, align)
        else:
            self._place_num_in_part(section.footer, align)
        self.set_page_number_format(section, fmt, start)
