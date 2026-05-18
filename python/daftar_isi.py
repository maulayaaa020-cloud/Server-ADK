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

def _ensure_toc_style(doc, level, use_dots, font='Times New Roman', size_pt=12, line_spacing=1.0):
    """
    Pastikan style 'TOC X' ada dengan font/size/spasi seragam + tab stop kanan.
    Menggunakan XML langsung agar semua script font (ascii/hAnsi/cs) dan bold
    ter-override secara eksplisit — menghindari inkonsistensi antar dokumen.
    """
    style_name = f'TOC {level}'
    try:
        style = doc.styles[style_name]
    except KeyError:
        from docx.enum.style import WD_STYLE_TYPE
        style = doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
        try:
            style.base_style = doc.styles['Normal']
        except Exception:
            pass

    # ── Paragraph format ─────────────────────────────────────────────────────
    pf = style.paragraph_format
    pf.left_indent  = Cm((level - 1) * 0.5)
    pf.space_after  = Pt(0)
    pf.line_spacing = line_spacing

    # ── Run properties: reset dan set ulang via XML ──────────────────────────
    # Ambil atau buat rPr di definisi style
    rPr = style.element.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        style.element.append(rPr)

    # Hapus semua formatting lama yang bisa konflik
    for tag in (qn('w:rFonts'), qn('w:sz'), qn('w:szCs'), qn('w:b'), qn('w:color')):
        old = rPr.find(tag)
        if old is not None:
            rPr.remove(old)

    # Font: set ascii + hAnsi + cs sekaligus
    fonts_el = OxmlElement('w:rFonts')
    fonts_el.set(qn('w:ascii'), font)
    fonts_el.set(qn('w:hAnsi'), font)
    fonts_el.set(qn('w:cs'),    font)
    rPr.insert(0, fonts_el)

    # Ukuran font (sz dan szCs untuk complex script)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(size_pt * 2))
    rPr.append(sz)
    szCs = OxmlElement('w:szCs')
    szCs.set(qn('w:val'), str(size_pt * 2))
    rPr.append(szCs)

    # Bold: H1 = bold, H2/H3 = eksplisit tidak bold (val=0)
    b = OxmlElement('w:b')
    if level > 1:
        b.set(qn('w:val'), '0')
    rPr.append(b)

    # ── Tab stop kanan dengan/tanpa leader titik ─────────────────────────────
    pPr = style.element.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        style.element.append(pPr)

    old_tabs = pPr.find(qn('w:tabs'))
    if old_tabs is not None:
        pPr.remove(old_tabs)

    tabs = OxmlElement('w:tabs')
    tab = OxmlElement('w:tab')
    tab.set(qn('w:val'), 'right')
    tab.set(qn('w:pos'), TAB_POS_TWIPS)
    tab.set(qn('w:leader'), 'dot' if use_dots else 'none')
    tabs.append(tab)
    pPr.append(tabs)


# ── Helper run ────────────────────────────────────────────────────────────────

def _make_run(text, bold=False, font='Times New Roman', size_pt=12):
    """Buat elemen <w:r> dengan formatting inline."""
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    fonts = OxmlElement('w:rFonts')
    fonts.set(qn('w:ascii'), font)
    fonts.set(qn('w:hAnsi'), font)
    rPr.append(fonts)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(size_pt * 2))
    rPr.append(sz)
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

    # ── Deteksi heading ──────────────────────────────────────────────────────
    headings = detect_headings(doc, max_level)

    if not headings:
        _fail("NO_HEADINGS_FOUND",
              "Tidak ditemukan heading di dokumen. Pastikan dokumen menggunakan "
              "style Heading 1/2/3 dari Word, atau teks dengan pola penomoran "
              "seperti '1.', '1.1', 'BAB I', atau teks ALL CAPS + bold.")

    # ── Terapkan outline level agar TOC field bisa mendeteksi semua heading ──
    for idx, lvl, _txt in headings:
        para = doc.paragraphs[idx]
        style_name = para.style.name if para.style else ''
        if not style_name.startswith('Heading '):
            apply_outline_level(para, lvl)

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
