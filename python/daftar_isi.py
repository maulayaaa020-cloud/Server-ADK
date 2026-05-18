"""
daftar_isi.py — Buat daftar isi otomatis dari dokumen Word.
Usage: python daftar_isi.py <input.docx> <output.docx> <kedalaman> <format_titik>

kedalaman:    H1 | H1+H2 | H1+H2+H3
format_titik: titik | tab
"""
import sys
import os
import json
import re
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Cm


# ── Konstanta ─────────────────────────────────────────────────────────────────

MAX_LEVEL_MAP = {'H1': 1, 'H1+H2': 2, 'H1+H2+H3': 3}

# Tab stop kanan untuk nomor halaman (twips: 1 cm ≈ 567 twips, ~15 cm = 8505)
TAB_POS_TWIPS = '8505'

_XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'


# ── Validasi file ─────────────────────────────────────────────────────────────

def _fail(code, msg):
    print(json.dumps({"status": "error", "code": code, "message": msg}))
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
        _fail("FILE_TOO_LARGE", f"Ukuran file ({size // (1024*1024)} MB) melebihi batas 30 MB.")
    try:
        with zipfile.ZipFile(path, 'r') as z:
            names = z.namelist()
            if 'word/document.xml' not in names:
                _fail("INVALID_DOCX", "File bukan dokumen Word (.docx) yang valid.")
            if 'word/vbaProject.bin' in names:
                _fail("MACRO_DETECTED",
                      "File mengandung macro VBA. Simpan ulang sebagai .docx biasa.")
            total = sum(i.file_size for i in z.infolist())
            if total > 200 * 1024 * 1024:
                _fail("FILE_TOO_LARGE", "Konten file melebihi 200 MB.")
    except zipfile.BadZipFile:
        _fail("INVALID_DOCX", "File rusak atau bukan format .docx yang valid.")


# ── Deteksi heading ───────────────────────────────────────────────────────────

def _get_outline_level_from_xml(para):
    """Baca w:outlineLvl dari XML paragraf (0-based), kembalikan 1-based."""
    pPr = para._p.find(qn('w:pPr'))
    if pPr is None:
        return None
    el = pPr.find(qn('w:outlineLvl'))
    if el is None:
        return None
    val = el.get(qn('w:val'))
    if val is None:
        return None
    try:
        lvl = int(val)
        return lvl + 1 if lvl < 9 else None
    except ValueError:
        return None


def _get_outline_level_from_style(para):
    """
    Naik ke rantai parent style dan cek w:outlineLvl di definisi style.
    Berguna untuk style kustom yang menyimpan outlineLvl di sana.
    """
    style = para.style
    while style:
        pPr = style.element.find(qn('w:pPr'))
        if pPr is not None:
            el = pPr.find(qn('w:outlineLvl'))
            if el is not None:
                try:
                    lvl = int(el.get(qn('w:val'), 9))
                    if lvl < 9:
                        return lvl + 1
                except (ValueError, TypeError):
                    pass
        style = style.base_style
    return None


def _get_num_level(para):
    """
    Baca level dari w:numPr (multilevel list Word), kembalikan ilvl 0-based atau None.
    numId=0 berarti penghapusan numbering, diabaikan.
    """
    pPr = para._p.find(qn('w:pPr'))
    if pPr is None:
        return None
    numPr = pPr.find(qn('w:numPr'))
    if numPr is None:
        return None
    numId_el = numPr.find(qn('w:numId'))
    if numId_el is not None:
        try:
            if int(numId_el.get(qn('w:val'), 1)) == 0:
                return None
        except (ValueError, TypeError):
            pass
    ilvl_el = numPr.find(qn('w:ilvl'))
    if ilvl_el is None:
        return None
    try:
        return int(ilvl_el.get(qn('w:val'), 0))
    except (ValueError, TypeError):
        return None


def _get_dominant_font_size(para):
    """Kembalikan ukuran font terbesar (dalam pt) dari run paragraf, atau None."""
    sizes = [r.font.size for r in para.runs if r.font.size]
    if not sizes:
        return None
    try:
        from docx.shared import Pt
        return max(sizes).pt
    except Exception:
        return None


_FRONT_MATTER_RE = re.compile(
    r'^(KATA\s+PENGANTAR|PRAKATA|UCAPAN\s+TERIMA\s+KASIH|SANWACANA|'
    r'ABSTRAK|ABSTRACT|INTISARI|SARI|RINGKASAN|SUMMARY|'
    r'DAFTAR\s+(ISI|GAMBAR|TABEL|LAMPIRAN|SINGKATAN|SIMBOL|NOTASI|'
    r'ARTI\s+LAMBANG|PERSAMAAN|PUSTAKA)|'
    r'LEMBAR\s+(PERSETUJUAN|PENGESAHAN|PERNYATAAN\S*|ORISINALITAS)|'
    r'HALAMAN\s+(PERSETUJUAN\S*|PENGESAHAN|PERNYATAAN\S*|PERSEMBAHAN|'
    r'JUDUL|COVER|ORISINALITAS)|'
    r'PERNYATAAN\s+(KEASLIAN|ORISINALITAS)|'
    r'MOTTO|PERSEMBAHAN|REFERENCES?|BIBLIOGRAPHY|DAFTAR\s+PUSTAKA|'
    r'HALAMAN\s+PERSETUJUAN\s+PEMBIMBING|HALAMAN\s+PERNYATAAN\s+KEASLIAN)$',
    re.IGNORECASE
)

_CUSTOM_H1_STYLES = {'H1', 'Heading Bab', 'Judul Bab', 'BAB', 'Bab', 'Judul'}
_CUSTOM_H2_STYLES = {'H2', 'Heading Sub', 'Sub Judul', 'Sub Bab', 'Sub-Bab'}
_CUSTOM_H3_STYLES = {'H3', 'Sub Sub Judul'}

_STYLE_KW_H1 = re.compile(
    r'(judul\s*(bab|utama|1)?|^bab$|heading\s*1?|chapter|title)$', re.IGNORECASE)
_STYLE_KW_H2 = re.compile(
    r'(sub[\s\-]?(judul|bab|heading)?[\s\-]?1?|heading\s*2|sub\s*title)$', re.IGNORECASE)
_STYLE_KW_H3 = re.compile(
    r'(sub[\s\-]?sub|heading\s*3|sub[\s\-]?(judul|bab)[\s\-]?2)$', re.IGNORECASE)


def get_para_level(para):
    """
    Kembalikan level heading (1-3) atau None.

    Patokan utama (berurutan):
      1. Pola BAB  → H1
      2. Pola numbering: 1.1.1→H3 | 1.1→H2 | 1.→H1
         (termasuk variasi spasi: "3. 10 Judul" → di-normalize dulu)
      3. Pola huruf: A.1→H3 | I.A→H2 | A.→H2
      4. Word numbering (numPr ilvl 1-2) → H2/H3
         (hanya jika style juga merupakan style heading, bukan List Paragraph)
      5. outlineLvl dari XML paragraf
      5b. outlineLvl dari definisi style (naik ke parent) — tangani style kustom H1/H2/H3
      6. Front matter keyword → H1
      7. Heading style standar Word — hanya untuk teks pendek (≤80 char)

    TIDAK dideteksi:
      - Teks panjang (> 150 char) → selalu diabaikan
      - Font size / style name kustom / ALL CAPS tanpa pola → dihapus
        (terlalu banyak false positive dari judul cover dan paragraf body)
    """
    text = para.text.strip()
    if not text or len(text) > 150:
        return None

    # Normalisasi penomoran dengan spasi ekstra setelah titik:
    # "3. 10 Judul" → "3.10 Judul", "1. 4 Judul" → "1.4 Judul"
    # Ini menangani kasus user mengetik nomor dengan spasi sebelum angka sub.
    text_norm = re.sub(r'^(\d+)\.\s+(\d)', r'\1.\2', text)

    # 1. Pola BAB → H1
    if re.match(r'^\s*BAB\s+[IVXivx\d]+\b', text, re.IGNORECASE):
        return 1

    # 2. Pola numbering bertitik (wajib titik agar tidak false-positive)
    #    Gunakan text_norm agar "3. 10 Judul" → "3.10" terdeteksi sebagai H2
    if re.match(r'^\d+\.\d+\.\d+\.?\s+\S', text_norm):
        return 3
    if re.match(r'^\d+\.\d+\.?\s+\S', text_norm):
        return 2
    if re.match(r'^\d+\.\s+\S', text):
        return 1

    # 3. Pola huruf dan kombinasi
    if re.match(r'^[A-Z]\.\d+\.?\s+\S', text):     # A.1 / B.2 → H3
        return 3
    if re.match(r'^[IVX]+\.[A-Z]\.?\s+\S', text):   # I.A / II.B → H2
        return 2
    if re.match(r'^[A-Z]\.\s+\S', text):             # A. / B. → H2
        return 2

    # 4. Word multilevel numbering (numPr) ilvl 1→H2, 2→H3
    #    Hanya berlaku jika style paragraf juga merupakan style heading,
    #    bukan 'List Paragraph' / 'List Bullet' / 'List Number' (list biasa).
    _LIST_STYLES = {'List Paragraph', 'List Bullet', 'List Number',
                    'List Bullet 2', 'List Bullet 3',
                    'List Number 2', 'List Number 3'}
    num_lvl    = _get_num_level(para)
    style_name = para.style.name if para.style else ''
    if num_lvl is not None and 1 <= num_lvl <= 2:
        if style_name not in _LIST_STYLES:
            return num_lvl + 1

    # 5. outlineLvl dari XML paragraf
    lvl = _get_outline_level_from_xml(para)
    if lvl is not None:
        return lvl

    # 5b. outlineLvl dari definisi style (naik ke parent style).
    #     Menangani style kustom seperti 'H1', 'H2', 'H3' yang tidak punya
    #     outlineLvl di paragraf tetapi punya di definisi style-nya.
    lvl = _get_outline_level_from_style(para)
    if lvl is not None:
        return lvl

    # 6. Front matter keyword (KATA PENGANTAR, ABSTRAK, DAFTAR ISI, dst.)
    if _FRONT_MATTER_RE.match(text):
        return 1

    # 7. Heading style standar Word — HANYA teks pendek
    #    Paragraf panjang ber-Heading (kecelakaan formatting) diabaikan.
    #    Cover / judul skripsi biasanya pakai style Normal, bukan Heading → aman.
    if len(text) <= 80:
        if style_name.startswith('Heading '):
            try:
                return int(style_name.split()[-1])
            except ValueError:
                pass

    return None


def _dedup_bab_clusters(results, max_gap=10):
    """
    Hapus BAB palsu dari bagian sistematika / daftar isi pengantar.

    Pola sistematika: penulis menyebut "BAB I ..., BAB II ..., BAB III ..."
    berurutan dalam satu paragraf area → jarak antar BAB sangat kecil dan
    tidak ada H2/H3 di antara mereka.

    Algoritma: jika dua entry BAB berurutan berjarak ≤ max_gap paragraf DAN
    tidak ada heading H2/H3 di antara keduanya dalam results, entry pertama
    kemungkinan besar palsu → hapus. Entry TERAKHIR dalam rantai tersebut
    selalu dipertahankan (itu BAB asli yang memulai konten).
    """
    _bab = re.compile(r'^\s*BAB\s+[IVXivx\d]+\b', re.IGNORECASE)
    to_remove = set()

    for k in range(len(results) - 1):
        pi1, lvl1, txt1 = results[k]
        pi2, lvl2, txt2 = results[k + 1]

        if lvl1 != 1 or not _bab.match(txt1):
            continue
        if lvl2 != 1 or not _bab.match(txt2):
            continue

        gap = pi2 - pi1
        if gap > max_gap:
            continue

        # Cek apakah ada H2/H3 di antara k dan k+1 dalam results
        has_sub = any(lvl >= 2 for _, lvl, _ in results[k + 1: k + 1])
        # (selalu False karena slice kosong — benar, kita cek gap bukan results)
        has_sub = any(
            lvl >= 2
            for j in range(k + 1, len(results))
            if results[j][0] < pi2
            for _, lvl, _ in [results[j]]
        )
        if not has_sub:
            to_remove.add(k)  # hapus entry PERTAMA, pertahankan entry berikutnya

    return [entry for i, entry in enumerate(results) if i not in to_remove]


def detect_headings(doc, max_level):
    """Kembalikan list of (para_index, level, text) untuk semua heading ≤ max_level."""
    paras   = doc.paragraphs
    results = []
    for i, para in enumerate(paras):
        lvl = get_para_level(para)
        if lvl is None or lvl > max_level:
            continue
        text = para.text.strip()
        # Jika H1 hanya berisi "BAB X" tanpa nama bab (dua paragraf terpisah),
        # gabungkan dengan paragraf ALL CAPS berikutnya sebagai nama bab.
        if lvl == 1 and re.match(r'^\s*BAB\s+[IVXivx\d]+\s*$', text, re.IGNORECASE):
            for j in range(i + 1, min(i + 3, len(paras))):
                nt = paras[j].text.strip()
                if nt and nt == nt.upper() and len(nt) > 3:
                    text = text + ' ' + nt
                    break
        results.append((i, lvl, text))

    results = _dedup_bab_clusters(results)
    return results


# ── Bold normalization untuk heading ─────────────────────────────────────────

def _strip_inline_bold_only(para):
    """Strip w:b/w:bCs dari pPr/rPr dan setiap run. TIDAK menambah w:b val=0."""
    BOLD_TAGS = (qn('w:b'), qn('w:bCs'))
    pPr = para._p.find(qn('w:pPr'))
    if pPr is not None:
        rPr = pPr.find(qn('w:rPr'))
        if rPr is not None:
            for tag in BOLD_TAGS:
                el = rPr.find(tag)
                if el is not None:
                    rPr.remove(el)
    for r_el in para._p.findall(qn('w:r')):
        rPr = r_el.find(qn('w:rPr'))
        if rPr is None:
            continue
        for tag in BOLD_TAGS:
            el = rPr.find(tag)
            if el is not None:
                rPr.remove(el)


def _set_runs_bold(para):
    """Tambahkan w:b/w:bCs ke semua runs (inline bold eksplisit)."""
    for r_el in para._p.findall(qn('w:r')):
        rPr = r_el.find(qn('w:rPr'))
        if rPr is None:
            rPr = OxmlElement('w:rPr')
            r_el.insert(0, rPr)
        for tag in (qn('w:b'), qn('w:bCs')):
            el = rPr.find(tag)
            if el is not None:
                rPr.remove(el)
        rPr.append(OxmlElement('w:b'))
        rPr.append(OxmlElement('w:bCs'))


def _normalize_heading_bold(para, level):
    """
    Normalisasi bold heading di body dan TOC.

    H1: tambahkan inline bold → body bold ✓, TOC bold ✓.
    H2/H3 + heading style (standard maupun custom H2/H3):
        strip inline bold saja → bold dari style, TOC tidak bold ✓.
    H2/H3 tanpa heading style: tambahkan inline bold → body bold ✓,
        TOC mungkin juga bold (limitation untuk non-standard style).
    """
    style_name = para.style.name if para.style else ''
    is_heading_style = (
        style_name.startswith('Heading ')
        or style_name in _CUSTOM_H1_STYLES
        or style_name in _CUSTOM_H2_STYLES
        or style_name in _CUSTOM_H3_STYLES
    )
    if level == 1:
        _set_runs_bold(para)
    else:
        _strip_inline_bold_only(para)
        if not is_heading_style:
            _set_runs_bold(para)


# ── Terapkan outline level ─────────────────────────────────────────────────────

def apply_outline_level(para, level):
    """
    Tambahkan w:outlineLvl ke paragraf tanpa mengubah tampilan visual.
    level: 1-based. w:outlineLvl adalah 0-based.
    """
    pPr = para._p.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        para._p.insert(0, pPr)

    existing = pPr.find(qn('w:outlineLvl'))
    if existing is not None:
        pPr.remove(existing)

    el = OxmlElement('w:outlineLvl')
    el.set(qn('w:val'), str(level - 1))
    pPr.append(el)


# ── TOC styles ────────────────────────────────────────────────────────────────

def _apply_rPr_font_to_style_element(style_el, font, size_pt, bold, is_char_style=False):
    """
    Terapkan font/size/bold ke elemen style XML secara menyeluruh.
    Menghapus semua theme-based font references (asciiTheme, hAnsiTheme, cstheme)
    yang bisa override font eksplisit saat Word update TOC field.

    is_char_style=True: hapus juga sz/color/kern agar tidak ada residual
    dari Heading X Char yang menimpa TOC entry saat dirender.
    """
    rPr = style_el.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        style_el.append(rPr)

    # Hapus tag yang bisa konflik
    remove_tags = [qn('w:rFonts'), qn('w:sz'), qn('w:szCs'), qn('w:b'), qn('w:bCs')]
    if is_char_style:
        # Untuk char style, hapus juga color/kern/ligatures agar tidak
        # ada efek visual dari Heading Char yang ikut masuk ke TOC entry
        remove_tags += [qn('w:color'), qn('w:kern'), qn('w:lang')]
        # Hapus themeColor/themeShade attribute dari color jika ada
    for tag in remove_tags:
        old = rPr.find(tag)
        if old is not None:
            rPr.remove(old)

    # Font: set ascii + hAnsi + cs secara eksplisit (TANPA theme reference)
    # Theme reference seperti asciiTheme="majorHAnsi" akan di-resolve oleh Word
    # menjadi Calibri Light / Cambria, BUKAN Times New Roman.
    fonts_el = OxmlElement('w:rFonts')
    fonts_el.set(qn('w:ascii'),  font)
    fonts_el.set(qn('w:hAnsi'),  font)
    fonts_el.set(qn('w:cs'),     font)
    fonts_el.set(qn('w:eastAsia'), font)
    rPr.insert(0, fonts_el)

    # Ukuran font
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(size_pt * 2))
    rPr.append(sz)
    szCs = OxmlElement('w:szCs')
    szCs.set(qn('w:val'), str(size_pt * 2))
    rPr.append(szCs)

    # Bold
    b = OxmlElement('w:b')
    if not bold:
        b.set(qn('w:val'), '0')
    rPr.append(b)
    bCs = OxmlElement('w:bCs')
    if not bold:
        bCs.set(qn('w:val'), '0')
    rPr.append(bCs)


def _fix_heading_char_styles(doc, font, size_pt):
    """
    Override Heading X Char styles agar font konsisten dan bold terkontrol.

    - Heading 1 Char: bold=True (TOC H1 harus bold)
    - Heading 2/3+ Char: bold=False (TOC H2/H3 tidak bold)

    Saat Word update TOC field, ia menaruh w:rStyle="HeadingXChar" di setiap run
    TOC entry. Style ini bisa override font TOC style jika tidak di-fix.
    """
    styles_root = doc.styles.element

    heading_char_ids = {
        'Heading1Char', 'Heading2Char', 'Heading3Char',
        'Heading4Char', 'Heading5Char', 'Heading6Char',
        'Heading7Char', 'Heading8Char', 'Heading9Char',
        'H1Char', 'H2Char', 'H3Char',
    }
    h1_char_ids = {'Heading1Char', 'H1Char'}

    for style_el in styles_root.findall(qn('w:style')):
        stype = style_el.get(qn('w:type'), '')
        if stype != 'character':
            continue
        style_id = style_el.get(qn('w:styleId'), '')
        sname_el = style_el.find(qn('w:name'))
        sname = sname_el.get(qn('w:val'), '') if sname_el is not None else ''

        is_heading_char = (
            style_id in heading_char_ids
            or ('Heading' in sname and 'Char' in sname)
            or (style_id.startswith('H') and style_id.endswith('Char')
                and style_id[1:-4].isdigit())
        )
        if not is_heading_char:
            continue

        # H1 Char: bold=True; semua lainnya: bold=False
        is_h1_char = (
            style_id in h1_char_ids
            or sname.lower() in ('heading 1 char', 'heading1char')
        )
        _apply_rPr_font_to_style_element(
            style_el, font=font, size_pt=size_pt, bold=is_h1_char, is_char_style=True
        )


def _ensure_toc_style(doc, level, use_dots, font='Times New Roman', size_pt=12, line_spacing=1.0):
    """
    Pastikan style 'TOC X' ada dengan font/size/spasi seragam + tab stop kanan.

    Menangani DUA versi style yang mungkin ada di dokumen Word Indonesia:
      - 'TOC 1' / 'TOC 2' / 'TOC 3'  (Pascal case — dipakai saat Word update field)
      - 'toc 1' / 'toc 2' / 'toc 3'  (lowercase — style bawaan Word Indonesia)

    Kedua versi harus punya font yang sama. Jika 'toc X' lowercase tidak punya
    rPr (seperti yang ditemukan di dokumen sumber), Word akan fallback ke Normal
    style yang fontnya Calibri — menyebabkan inkonsistensi.

    Menggunakan XML langsung agar semua script font (ascii/hAnsi/cs/eastAsia)
    ter-override secara eksplisit tanpa theme reference.
    """
    from docx.enum.style import WD_STYLE_TYPE

    # Tangani kedua versi nama style (Pascal dan lowercase)
    style_names_to_fix = [f'TOC {level}', f'toc {level}']

    for style_name in style_names_to_fix:
        try:
            style = doc.styles[style_name]
        except KeyError:
            if style_name == f'TOC {level}':
                # Hanya buat style TOC jika belum ada (versi Pascal case adalah utama)
                style = doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
                try:
                    style.base_style = doc.styles['Normal']
                except Exception:
                    pass
            else:
                # Versi lowercase opsional — skip jika tidak ada
                continue

        # ── Paragraph format ─────────────────────────────────────────────────
        pf = style.paragraph_format
        pf.space_after  = Pt(0)
        pf.line_spacing = line_spacing

        # Indentasi per level (dari analisis referensi)
        # TOC1: left=0
        # TOC2: left=284t, tidak ada hanging
        # TOC3: left=1560t, hanging=709t (teks mulai di 1560-709=851t, sama dgn TOC2)
        from docx.shared import Twips as _Twips
        IND = {
            1: (0,    0),
            2: (284,  0),
            3: (1560, 709),
        }
        left_t, hang_t = IND.get(level, (284 * (level - 1), 0))
        pf.left_indent       = _Twips(left_t)
        pf.first_line_indent = _Twips(-hang_t) if hang_t else None

        # ── Run properties via XML ────────────────────────────────────────────
        _apply_rPr_font_to_style_element(
            style.element,
            font=font,
            size_pt=size_pt,
            bold=(level == 1),
            is_char_style=False,
        )

        # ── Tab stops ─────────────────────────────────────────────────────────
        # Hitung posisi tab kanan (nomor halaman) dari margin dokumen
        try:
            sec = doc.sections[0]
            right_tab = int((sec.page_width - sec.left_margin - sec.right_margin).twips)
        except Exception:
            right_tab = 7927  # fallback ~14cm

        # TOC2: DUA tab stop — tab kiri 851t (alignment teks) + tab kanan (halaman)
        # TOC3: hanya tab kanan (alignment diatur oleh left+hanging indent)
        # TOC1: hanya tab kanan
        LEFT_TAB_TWIPS = {2: 851}  # dari analisis referensi terbaru

        pPr = style.element.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr')
            style.element.append(pPr)

        old_tabs = pPr.find(qn('w:tabs'))
        if old_tabs is not None:
            pPr.remove(old_tabs)

        tabs_el = OxmlElement('w:tabs')

        if level in LEFT_TAB_TWIPS:
            left_tab = OxmlElement('w:tab')
            left_tab.set(qn('w:val'),    'left')
            left_tab.set(qn('w:pos'),    str(LEFT_TAB_TWIPS[level]))
            left_tab.set(qn('w:leader'), 'none')
            tabs_el.append(left_tab)

        right_tab_el = OxmlElement('w:tab')
        right_tab_el.set(qn('w:val'),    'right')
        right_tab_el.set(qn('w:pos'),    str(right_tab))
        right_tab_el.set(qn('w:leader'), 'dot' if use_dots else 'none')
        tabs_el.append(right_tab_el)

        pPr.append(tabs_el)


# ── Helper run ────────────────────────────────────────────────────────────────

def _make_run(text, bold=False, font='Times New Roman', size_pt=12):
    """Buat elemen <w:r> dengan formatting inline (semua script font di-set eksplisit)."""
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    fonts = OxmlElement('w:rFonts')
    fonts.set(qn('w:ascii'),   font)
    fonts.set(qn('w:hAnsi'),   font)
    fonts.set(qn('w:cs'),      font)
    fonts.set(qn('w:eastAsia'), font)
    rPr.append(fonts)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(size_pt * 2))
    rPr.append(sz)
    szCs = OxmlElement('w:szCs')
    szCs.set(qn('w:val'), str(size_pt * 2))
    rPr.append(szCs)
    if bold:
        rPr.append(OxmlElement('w:b'))
    r.append(rPr)
    t = OxmlElement('w:t')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text
    r.append(t)
    return r


def _make_entry_pPr(level, use_dots):
    """Buat pPr untuk satu entri TOC."""
    pPr = OxmlElement('w:pPr')
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:after'), '0')
    pPr.append(spacing)
    if level > 1:
        ind = OxmlElement('w:ind')
        ind.set(qn('w:left'), str((level - 1) * 360))
        pPr.append(ind)
    tabs = OxmlElement('w:tabs')
    tab = OxmlElement('w:tab')
    tab.set(qn('w:val'), 'right')
    tab.set(qn('w:pos'), TAB_POS_TWIPS)
    if use_dots:
        tab.set(qn('w:leader'), 'dot')
    tabs.append(tab)
    pPr.append(tabs)
    return pPr


def _make_toc_field_para(max_level):
    """
    Buat TOC field dengan dirty=true.
    Word otomatis hitung nomor halaman saat file dibuka (popup → klik Yes → selesai).
    Setelah user simpan file (Ctrl+S), popup tidak muncul lagi.
    """
    p = OxmlElement('w:p')

    r_begin = OxmlElement('w:r')
    fc = OxmlElement('w:fldChar')
    fc.set(qn('w:fldCharType'), 'begin')
    fc.set(qn('w:dirty'), 'true')
    r_begin.append(fc)
    p.append(r_begin)

    r_instr = OxmlElement('w:r')
    instr = OxmlElement('w:instrText')
    instr.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    instr.text = f' TOC \\o "1-{max_level}" \\z \\u '
    r_instr.append(instr)
    p.append(r_instr)

    r_sep = OxmlElement('w:r')
    fc_sep = OxmlElement('w:fldChar')
    fc_sep.set(qn('w:fldCharType'), 'separate')
    r_sep.append(fc_sep)
    p.append(r_sep)

    r_end = OxmlElement('w:r')
    fc_end = OxmlElement('w:fldChar')
    fc_end.set(qn('w:fldCharType'), 'end')
    r_end.append(fc_end)
    p.append(r_end)

    return p


# ── Static TOC (tanpa field) ──────────────────────────────────────────────────

def _add_paragraph_bookmark(para, bk_id, bk_name):
    """Sisipkan bookmark start/end di paragraf heading untuk dirujuk PAGEREF."""
    bkStart = OxmlElement('w:bookmarkStart')
    bkStart.set(qn('w:id'),   str(bk_id))
    bkStart.set(qn('w:name'), bk_name)
    bkEnd = OxmlElement('w:bookmarkEnd')
    bkEnd.set(qn('w:id'), str(bk_id))
    para._p.insert(0, bkStart)
    para._p.append(bkEnd)


def _build_run_rPr(font, size_pt, bold):
    """Buat w:rPr dengan font/size/bold eksplisit — tanpa theme reference."""
    rPr = OxmlElement('w:rPr')
    fonts_el = OxmlElement('w:rFonts')
    fonts_el.set(qn('w:ascii'),    font)
    fonts_el.set(qn('w:hAnsi'),    font)
    fonts_el.set(qn('w:cs'),       font)
    fonts_el.set(qn('w:eastAsia'), font)
    rPr.append(fonts_el)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(size_pt * 2))
    rPr.append(sz)
    szCs = OxmlElement('w:szCs')
    szCs.set(qn('w:val'), str(size_pt * 2))
    rPr.append(szCs)
    b = OxmlElement('w:b')
    if not bold:
        b.set(qn('w:val'), '0')
    rPr.append(b)
    bCs = OxmlElement('w:bCs')
    if not bold:
        bCs.set(qn('w:val'), '0')
    rPr.append(bCs)
    return rPr


def _make_static_toc_para(text, level, bk_name, font, size_pt, use_dots, right_tab_pos):
    """
    Buat satu paragraf TOC statis dengan PAGEREF untuk nomor halaman.

    Keuntungan vs TOC field:
    - Formatting (font, bold) sepenuhnya dikontrol kita via rPr inline
    - Tidak ada warisan bold/font dari paragraf sumber (DAF2 problem)
    - Nomor halaman tetap dinamis via PAGEREF — diupdate Word saat dibuka
    """
    import copy
    p       = OxmlElement('w:p')
    is_bold = (level == 1)

    # ── pPr ──────────────────────────────────────────────────────────────────
    pPr = OxmlElement('w:pPr')

    pStyle = OxmlElement('w:pStyle')
    pStyle.set(qn('w:val'), f'TOC{level}')
    pPr.append(pStyle)

    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:after'), '0')
    spacing.set(qn('w:line'), '360')
    spacing.set(qn('w:lineRule'), 'auto')
    pPr.append(spacing)

    IND = {1: (0, 0), 2: (284, 0), 3: (1560, 709)}
    left_t, hang_t = IND.get(level, (284 * (level - 1), 0))
    if left_t or hang_t:
        ind = OxmlElement('w:ind')
        if left_t:  ind.set(qn('w:left'),    str(left_t))
        if hang_t:  ind.set(qn('w:hanging'), str(hang_t))
        pPr.append(ind)

    LEFT_TAB = {2: 851}
    tabs_el = OxmlElement('w:tabs')
    if level in LEFT_TAB:
        lt = OxmlElement('w:tab')
        lt.set(qn('w:val'), 'left')
        lt.set(qn('w:pos'), str(LEFT_TAB[level]))
        lt.set(qn('w:leader'), 'none')
        tabs_el.append(lt)
    rt = OxmlElement('w:tab')
    rt.set(qn('w:val'), 'right')
    rt.set(qn('w:pos'), str(right_tab_pos))
    rt.set(qn('w:leader'), 'dot' if use_dots else 'none')
    tabs_el.append(rt)
    pPr.append(tabs_el)

    p.append(pPr)

    # ── Teks heading ──────────────────────────────────────────────────────────
    r_text = OxmlElement('w:r')
    r_text.append(_build_run_rPr(font, size_pt, is_bold))
    t = OxmlElement('w:t')
    t.set(_XML_SPACE, 'preserve')
    t.text = text.replace('\n', ' ')
    r_text.append(t)
    p.append(r_text)

    # ── Tab ───────────────────────────────────────────────────────────────────
    r_tab = OxmlElement('w:r')
    r_tab.append(_build_run_rPr(font, size_pt, is_bold))
    r_tab.append(OxmlElement('w:tab'))
    p.append(r_tab)

    # ── PAGEREF field ─────────────────────────────────────────────────────────
    base_rPr = _build_run_rPr(font, size_pt, is_bold)

    r_begin = OxmlElement('w:r')
    r_begin.append(copy.deepcopy(base_rPr))
    fc = OxmlElement('w:fldChar')
    fc.set(qn('w:fldCharType'), 'begin')
    r_begin.append(fc)
    p.append(r_begin)

    r_instr = OxmlElement('w:r')
    r_instr.append(copy.deepcopy(base_rPr))
    instr = OxmlElement('w:instrText')
    instr.set(_XML_SPACE, 'preserve')
    instr.text = f' PAGEREF {bk_name} \\h '
    r_instr.append(instr)
    p.append(r_instr)

    r_sep = OxmlElement('w:r')
    r_sep.append(copy.deepcopy(base_rPr))
    fc_sep = OxmlElement('w:fldChar')
    fc_sep.set(qn('w:fldCharType'), 'separate')
    r_sep.append(fc_sep)
    p.append(r_sep)

    r_ph = OxmlElement('w:r')
    r_ph.append(copy.deepcopy(base_rPr))
    t_ph = OxmlElement('w:t')
    t_ph.text = '1'
    r_ph.append(t_ph)
    p.append(r_ph)

    r_end = OxmlElement('w:r')
    r_end.append(copy.deepcopy(base_rPr))
    fc_end = OxmlElement('w:fldChar')
    fc_end.set(qn('w:fldCharType'), 'end')
    r_end.append(fc_end)
    p.append(r_end)

    return p


# ── Helper page break ─────────────────────────────────────────────────────────

def _has_inline_page_break(para):
    for br in para._p.findall('.//' + qn('w:br')):
        if br.get(qn('w:type')) == 'page':
            return True
    return False


def _remove_inline_page_break(para):
    for r_elem in list(para._p.findall(qn('w:r'))):
        br = r_elem.find(qn('w:br'))
        if br is not None and br.get(qn('w:type')) == 'page':
            para._p.remove(r_elem)


def _make_page_break_para():
    p = OxmlElement('w:p')
    r = OxmlElement('w:r')
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    r.append(br)
    p.append(r)
    return p


# ── Cari posisi DAFTAR ISI ────────────────────────────────────────────────────

def find_daftar_isi_idx(doc):
    """Kembalikan index paragraf 'DAFTAR ISI', atau None jika tidak ada."""
    for i, para in enumerate(doc.paragraphs):
        if re.match(r'^\s*daftar\s+isi\s*$', para.text.strip(), re.IGNORECASE):
            return i
    return None


# ── Auto-update fields on open ────────────────────────────────────────────────

def enable_auto_update_fields(doc):
    """Tambah w:updateFields=true ke settings.xml agar Word langsung update tanpa prompt."""
    settings = doc.settings.element
    existing = settings.find(qn('w:updateFields'))
    if existing is not None:
        settings.remove(existing)
    el = OxmlElement('w:updateFields')
    el.set(qn('w:val'), 'true')
    settings.append(el)


# ── Sisipkan setelah paragraf ──────────────────────────────────────────────────

def insert_element_after(doc, para_idx, new_elem):
    """Sisipkan elemen XML setelah paragraf pada index para_idx."""
    paras = doc.paragraphs
    ref = paras[para_idx]._element
    ref.addnext(new_elem)


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 5:
        _fail("MISSING_ARGUMENTS",
              "Usage: python daftar_isi.py <input> <output> <kedalaman> <format_titik> [font] [size] [space]")

    input_file   = sys.argv[1]
    output_file  = sys.argv[2]
    kedalaman    = sys.argv[3]   # H1 | H1+H2 | H1+H2+H3
    format_titik = sys.argv[4]   # titik | tab | kecualikan_bab
    font         = sys.argv[5] if len(sys.argv) > 5 else 'Times New Roman'
    size_pt      = int(sys.argv[6]) if len(sys.argv) > 6 else 12
    try:
        line_spacing = float(sys.argv[7]) if len(sys.argv) > 7 else 1.0
    except (ValueError, TypeError):
        line_spacing = 1.0

    max_level = MAX_LEVEL_MAP.get(kedalaman, 1)
    use_dots  = format_titik in ('titik', 'kecualikan_bab')

    validate_file(input_file)

    try:
        doc = Document(input_file)
    except Exception as e:
        _fail("FILE_READ_ERROR", f"Gagal membuka file: {e}")

    # ── Seragamkan style TOC agar font/size/spasi konsisten ─────────────────
    for lvl in range(1, max_level + 1):
        _ensure_toc_style(doc, lvl, use_dots, font=font, size_pt=size_pt, line_spacing=line_spacing)

    # ── Fix Heading X Char styles (root cause inkonsistensi font) ────────────
    # Saat Word update TOC field, ia menaruh rStyle="Heading2Char" di setiap
    # run paragraf TOC 2/3. Style karakter ini menggunakan theme font
    # (majorHAnsi = Calibri Light) yang OVERRIDE font Times New Roman dari
    # definisi style TOC. Solusi: samakan font Heading X Char dengan TOC font.
    _fix_heading_char_styles(doc, font=font, size_pt=size_pt)

    # ── Deteksi heading ──────────────────────────────────────────────────────
    headings = detect_headings(doc, max_level)

    if not headings:
        _fail("NO_HEADINGS_FOUND",
              "Tidak ditemukan heading di dokumen. Pastikan dokumen menggunakan "
              "style Heading 1/2/3 dari Word, atau teks dengan pola penomoran "
              "seperti '1.', '1.1', 'BAB I', atau teks ALL CAPS + bold.")

    # ── Pastikan semua heading style (standard + custom) punya bold ───────────
    # Sehingga setelah inline bold di-strip, body heading masih tampil bold.
    _styles_to_bold = (
        [f'Heading {i}' for i in range(1, max_level + 1)]
        + list(_CUSTOM_H1_STYLES) + list(_CUSTOM_H2_STYLES) + list(_CUSTOM_H3_STYLES)
    )
    for sname in _styles_to_bold:
        try:
            style = doc.styles[sname]
            style_el = style.element
            rPr = style_el.find(qn('w:rPr'))
            if rPr is None:
                rPr = OxmlElement('w:rPr')
                style_el.append(rPr)
            for tag in (qn('w:b'), qn('w:bCs')):
                el = rPr.find(tag)
                if el is not None:
                    rPr.remove(el)
            rPr.append(OxmlElement('w:b'))
            rPr.append(OxmlElement('w:bCs'))
        except KeyError:
            pass

    # ── Pindahkan H2/H3 non-heading-style ke style H2/H3 ────────────────────
    # Tujuan: agar bold di body berasal dari style (bukan inline run),
    # sehingga TOC tidak mewarisi inline bold saat Word update field.
    # H2/H3 style sudah ada di dokumen dengan font/size sama (TNR 12pt).
    for idx, lvl, _txt in headings:
        if lvl < 2:
            continue
        para = doc.paragraphs[idx]
        style_name = para.style.name if para.style else ''
        is_heading_style = (
            style_name.startswith('Heading ')
            or style_name in _CUSTOM_H1_STYLES
            or style_name in _CUSTOM_H2_STYLES
            or style_name in _CUSTOM_H3_STYLES
        )
        if not is_heading_style:
            target = f'H{lvl}'
            try:
                para.style = doc.styles[target]
            except KeyError:
                pass  # style H2/H3 tidak ada di dokumen ini

    # ── Terapkan outline level agar TOC field bisa mendeteksi semua heading ──
    for idx, lvl, _txt in headings:
        para = doc.paragraphs[idx]
        style_name = para.style.name if para.style else ''
        if not style_name.startswith('Heading '):
            apply_outline_level(para, lvl)

    # ── Normalisasi bold: body selalu bold, TOC hanya H1 bold ────────────────
    for idx, lvl, _txt in headings:
        _normalize_heading_bold(doc.paragraphs[idx], lvl)

    # ── Aktifkan auto-update agar Word langsung hitung nomor halaman ─────────
    enable_auto_update_fields(doc)

    # ── Buat TOC field ───────────────────────────────────────────────────────
    toc_paras = [_make_toc_field_para(max_level)]

    # ── Sisipkan setelah "DAFTAR ISI" ────────────────────────────────────────
    daftar_idx = find_daftar_isi_idx(doc)
    if daftar_idx is not None:
        daftar_para = doc.paragraphs[daftar_idx]
        had_pb = _has_inline_page_break(daftar_para)
        if had_pb:
            _remove_inline_page_break(daftar_para)

        for new_p in reversed(toc_paras):
            insert_element_after(doc, daftar_idx, new_p)

        if had_pb:
            toc_paras[-1].addnext(_make_page_break_para())

        insert_pos_desc = f'setelah paragraf "DAFTAR ISI" (index {daftar_idx})'
    else:
        insert_at = 0
        for i, para in enumerate(doc.paragraphs):
            if para.text.strip():
                insert_at = i
                break
        for new_p in reversed(toc_paras):
            insert_element_after(doc, insert_at, new_p)
        insert_pos_desc = 'di awal dokumen (paragraf DAFTAR ISI tidak ditemukan)'

    # ── Simpan ───────────────────────────────────────────────────────────────
    try:
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        doc.save(output_file)
    except Exception as e:
        _fail("FILE_SAVE_ERROR", f"Gagal menyimpan file: {e}")

    print(json.dumps({
        "status":           "success",
        "kedalaman":        kedalaman,
        "format_titik":     format_titik,
        "heading_count":    len(headings),
        "insert_position":  insert_pos_desc,
        "headings_preview": [
            {"level": lvl, "text": txt[:60]}
            for _, lvl, txt in headings[:10]
        ],
    }))


if __name__ == '__main__':
    main()
