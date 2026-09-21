"""
Ubah dokumen .docx hasil render menjadi PDF.

Konverter ini membaca dokumen yang BARU SAJA dibuat docxtpl, lalu menyusun
ulang isinya ke PDF: kop surat di header Word, paragraf, dan tabel. Karena
semuanya dibaca dari dokumen hasil render, PDF otomatis mengikuti kalau
template Word diubah.

Catatan jujur: ini penyusunan ulang, bukan konversi piksel-per-piksel.
Konversi yang benar-benar identik membutuhkan LibreOffice atau Microsoft
Word, dan keduanya tidak tersedia di lingkungan serverless.
"""
import io
import zipfile
from xml.etree import ElementTree as ET

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.table import Table
from docx.text.paragraph import Paragraph
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.utils import ImageReader
from reportlab.platypus import KeepTogether, Paragraph as PdfParagraph
from reportlab.platypus import SimpleDocTemplate, Spacer, Table as PdfTable, TableStyle

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
}
EMU = 914400.0          # EMU per inci
PT = 72.0               # titik per inci

RATA = {
    WD_ALIGN_PARAGRAPH.CENTER: TA_CENTER,
    WD_ALIGN_PARAGRAPH.RIGHT: TA_RIGHT,
    WD_ALIGN_PARAGRAPH.JUSTIFY: TA_JUSTIFY,
}
RATA_XML = {'center': TA_CENTER, 'right': TA_RIGHT, 'both': TA_JUSTIFY, 'left': TA_LEFT}


# ---------------------------------------------------------------- font

def font_pdf(nama, tebal=False, miring=False):
    """Petakan nama font Word ke font bawaan PDF yang paling mendekati."""
    n = (nama or '').lower()
    if 'times' in n or 'serif' in n or 'georgia' in n:
        keluarga = 'Times'
        varian = {(0, 0): '-Roman', (1, 0): '-Bold', (0, 1): '-Italic', (1, 1): '-BoldItalic'}
    elif 'courier' in n or 'mono' in n or 'consol' in n:
        keluarga = 'Courier'
        varian = {(0, 0): '', (1, 0): '-Bold', (0, 1): '-Oblique', (1, 1): '-BoldOblique'}
    else:
        # Arial, Arial MT, Calibri, Helvetica, dan lain-lain -> Helvetica
        keluarga = 'Helvetica'
        varian = {(0, 0): '', (1, 0): '-Bold', (0, 1): '-Oblique', (1, 1): '-BoldOblique'}
    return keluarga + varian[(int(bool(tebal)), int(bool(miring)))]


def _escape(teks):
    return (teks or '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def _dari_gaya(paragraf, ambil):
    """Telusuri gaya paragraf dan induknya sampai menemukan nilai.

    Word menyimpan format judul di gaya (mis. 'Heading 1'), bukan di run,
    jadi membaca run saja membuat judul kehilangan ukuran dan ketebalannya.
    """
    gaya = getattr(paragraf, 'style', None)
    while gaya is not None:
        try:
            nilai = ambil(gaya)
        except AttributeError:
            nilai = None
        if nilai is not None:
            return nilai
        gaya = gaya.base_style
    return None


def _sifat(paragraf, bawaan_ukuran, bawaan_font):
    """Ukuran, nama font, tebal, dan perataan efektif sebuah paragraf."""
    ukuran = None
    nama = None
    tebal = None
    for run in paragraf.runs:
        if run.text.strip() == '':
            continue
        ukuran = ukuran or (run.font.size.pt if run.font.size else None)
        nama = nama or run.font.name
        if run.bold:
            tebal = True
    if ukuran is None:
        s = _dari_gaya(paragraf, lambda g: g.font.size)
        ukuran = s.pt if s else bawaan_ukuran
    if nama is None:
        nama = _dari_gaya(paragraf, lambda g: g.font.name) or bawaan_font
    if tebal is None:
        tebal = bool(_dari_gaya(paragraf, lambda g: g.font.bold))

    if paragraf.alignment is not None:
        rata = RATA.get(paragraf.alignment, TA_LEFT)
    else:
        a = _dari_gaya(paragraf, lambda g: g.paragraph_format.alignment)
        if a is not None:
            rata = RATA.get(a, TA_LEFT)
        else:
            rata = TA_CENTER if paragraf.style.name.startswith('Heading') else TA_LEFT
    return ukuran, nama, tebal, rata


def _inline(paragraf, tebal_bawaan=False):
    bagian = []
    for run in paragraf.runs:
        potongan = _escape(run.text)
        if not potongan:
            continue
        if run.bold and not tebal_bawaan:
            potongan = f'<b>{potongan}</b>'
        if run.italic:
            potongan = f'<i>{potongan}</i>'
        bagian.append(potongan)
    return ''.join(bagian) or _escape(paragraf.text)


def _gaya_pdf(nama, ukuran, font, tebal, rata, spasi_bawah):
    return ParagraphStyle(
        nama,
        fontName=font_pdf(font, tebal),
        fontSize=ukuran,
        leading=ukuran * 1.30,
        alignment=rata,
        spaceAfter=spasi_bawah,
    )


# ---------------------------------------------------------------- kop surat

def baca_kop(docx_bytes):
    """Ambil isi header Word sebagai daftar perintah gambar.

    Semua bentuk di header memakai koordinat relatif halaman, jadi bisa
    digambar ulang apa adanya di setiap halaman PDF.
    """
    perintah = []
    try:
        z = zipfile.ZipFile(io.BytesIO(docx_bytes))
        nama_header = [n for n in z.namelist() if n.startswith('word/header')]
        if not nama_header:
            return perintah
        header = sorted(nama_header)[0]

        rels = {}
        nama_rels = header.replace('word/', 'word/_rels/') + '.rels'
        if nama_rels in z.namelist():
            for rel in ET.fromstring(z.read(nama_rels)):
                rels[rel.get('Id')] = rel.get('Target')

        root = ET.fromstring(z.read(header))
        for anchor in root.iter(f'{{{NS["wp"]}}}anchor'):
            ext = anchor.find(f'{{{NS["wp"]}}}extent')
            if ext is None:
                continue
            lebar = int(ext.get('cx')) / EMU
            tinggi = int(ext.get('cy')) / EMU

            def offset(tag):
                p = anchor.find(f'{{{NS["wp"]}}}{tag}')
                if p is None:
                    return 0.0
                o = p.find(f'{{{NS["wp"]}}}posOffset')
                return int(o.text) / EMU if o is not None else 0.0

            x, y = offset('positionH'), offset('positionV')

            blip = anchor.find(f'.//{{{NS["a"]}}}blip')
            txbx = anchor.find(f'.//{{{NS["wps"]}}}txbx')

            if blip is not None:
                target = rels.get(blip.get(f'{{{NS["r"]}}}embed'))
                if not target:
                    continue
                jalur = 'word/' + target.replace('../', '')
                if jalur in z.namelist():
                    perintah.append(('gambar', x, y, lebar, tinggi, z.read(jalur)))
            elif txbx is not None:
                baris = []
                for p in txbx.iter(f'{{{NS["w"]}}}p'):
                    teks = ''.join(t.text or '' for t in p.iter(f'{{{NS["w"]}}}t'))
                    if not teks.strip():
                        continue
                    sz = p.find(f'.//{{{NS["w"]}}}sz')
                    rf = p.find(f'.//{{{NS["w"]}}}rFonts')
                    jc = p.find(f'.//{{{NS["w"]}}}jc')
                    b = p.find(f'.//{{{NS["w"]}}}b')
                    baris.append({
                        'teks': teks,
                        'ukuran': int(sz.get(f'{{{NS["w"]}}}val')) / 2 if sz is not None else None,
                        'font': rf.get(f'{{{NS["w"]}}}ascii') if rf is not None else None,
                        'rata': RATA_XML.get(jc.get(f'{{{NS["w"]}}}val') if jc is not None else '', TA_LEFT),
                        'tebal': b is not None,
                    })
                if baris:
                    perintah.append(('teks', x, y, lebar, tinggi, baris))
            elif tinggi <= 0.02:
                perintah.append(('garis', x, y, lebar, tinggi, None))
    except Exception:
        # Kop hanyalah hiasan; kalau gagal dibaca, dokumen tetap harus terbit.
        return []
    return perintah


def gambar_kop(canvas, perintah, tinggi_halaman):
    """Gambar kop di halaman saat ini."""
    for jenis, x, y, lebar, tinggi, data in perintah:
        X = x * PT
        Y = tinggi_halaman - (y + tinggi) * PT
        if jenis == 'gambar':
            try:
                canvas.drawImage(ImageReader(io.BytesIO(data)), X, Y,
                                 width=lebar * PT, height=tinggi * PT, mask='auto')
            except Exception:
                pass
        elif jenis == 'garis':
            canvas.setStrokeColor(colors.black)
            canvas.setLineWidth(1)
            canvas.line(X, Y, X + lebar * PT, Y)
        elif jenis == 'teks':
            atas = tinggi_halaman - y * PT
            for baris in data:
                ukuran = baris['ukuran'] or 11
                canvas.setFont(font_pdf(baris['font'], baris['tebal']), ukuran)
                atas -= ukuran * 1.18
                teks = baris['teks']
                # Potong kalau melebihi kotak, seperti Word memotong luapan.
                while canvas.stringWidth(teks, canvas._fontname, ukuran) > lebar * PT and len(teks) > 4:
                    teks = teks[:-1]
                if baris['rata'] == TA_CENTER:
                    canvas.drawCentredString(X + lebar * PT / 2, atas, teks)
                elif baris['rata'] == TA_RIGHT:
                    canvas.drawRightString(X + lebar * PT, atas, teks)
                else:
                    canvas.drawString(X, atas, teks)


# ---------------------------------------------------------------- isi

def _isi_blok(induk):
    """Hasilkan paragraf dan tabel sesuai urutan aslinya di dokumen."""
    for anak in induk.element.body.iterchildren():
        if anak.tag.endswith('}p'):
            yield Paragraph(anak, induk)
        elif anak.tag.endswith('}tbl'):
            yield Table(anak, induk)


def _sel_teks(sel, bawaan_ukuran, bawaan_font):
    bagian = []
    ukuran = bawaan_ukuran
    font = bawaan_font
    tebal = False
    rata = TA_LEFT
    for p in sel.paragraphs:
        potongan = _inline(p)
        if potongan:
            ukuran, font, tebal, rata = _sifat(p, bawaan_ukuran, bawaan_font)
            bagian.append(potongan)
    teks = '<br/>'.join(bagian)
    return PdfParagraph(teks, _gaya_pdf('sel', ukuran, font, tebal, rata, 0))


def _lebar_kolom(tabel):
    """Lebar kolom dalam titik, langsung dari definisi grid Word."""
    kolom = []
    for g in tabel._tbl.findall(f'.//{{{NS["w"]}}}gridCol'):
        w = g.get(f'{{{NS["w"]}}}w')
        if w:
            kolom.append(int(w) / 20.0)   # twip -> titik
    return kolom


def _bangun_tabel(tabel, lebar_maks, bawaan_ukuran, bawaan_font):
    kolom = len(tabel.columns)
    if not kolom:
        return None

    data = []
    perintah = [
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]

    for i, baris in enumerate(tabel.rows):
        sel_baris = list(baris.cells)
        terlihat, mulai = [], 0
        for j, sel in enumerate(sel_baris):
            if j > 0 and sel._tc is sel_baris[j - 1]._tc:
                continue
            if terlihat and j - 1 > mulai:
                perintah.append(('SPAN', (mulai, i), (j - 1, i)))
            terlihat.append(sel)
            mulai = j
        if mulai < len(sel_baris) - 1:
            perintah.append(('SPAN', (mulai, i), (len(sel_baris) - 1, i)))

        isi = [_sel_teks(s, bawaan_ukuran, bawaan_font) for s in terlihat]
        isi += [''] * (kolom - len(isi))
        data.append(isi)

        if i == 0 and len(terlihat) == 1:
            perintah.append(('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#d9d9d9')))

    lebar = _lebar_kolom(tabel)
    if len(lebar) != kolom or not sum(lebar):
        lebar = [lebar_maks / kolom] * kolom
    elif sum(lebar) > lebar_maks:
        skala = lebar_maks / sum(lebar)
        lebar = [w * skala for w in lebar]

    t = PdfTable(data, colWidths=lebar, hAlign='CENTER')
    t.setStyle(TableStyle(perintah))
    return t


def docx_ke_pdf(docx_bytes):
    """Terima byte .docx, balas byte PDF."""
    doc = Document(io.BytesIO(docx_bytes))
    sec = doc.sections[0]

    normal = doc.styles['Normal'].font
    bawaan_font = normal.name or 'Arial'
    bawaan_ukuran = normal.size.pt if normal.size else 11

    kiri = sec.left_margin.pt if sec.left_margin else 50
    kanan = sec.right_margin.pt if sec.right_margin else 50
    atas = sec.top_margin.pt if sec.top_margin else 50
    bawah = sec.bottom_margin.pt if sec.bottom_margin else 40

    kop = baca_kop(docx_bytes)

    buffer = io.BytesIO()
    pdf = SimpleDocTemplate(
        buffer, pagesize=A4,
        leftMargin=kiri, rightMargin=kanan, topMargin=atas, bottomMargin=bawah,
        title=doc.paragraphs[0].text if doc.paragraphs else 'Dokumen',
    )
    lebar_isi = A4[0] - kiri - kanan

    cerita = []
    for blok in _isi_blok(doc):
        if isinstance(blok, Paragraph):
            teks = _inline(blok)
            if not teks.strip():
                cerita.append(Spacer(1, 4))
                continue
            ukuran, font, tebal, rata = _sifat(blok, bawaan_ukuran, bawaan_font)
            cerita.append(PdfParagraph(teks, _gaya_pdf('p', ukuran, font, tebal, rata, 6)))
        else:
            t = _bangun_tabel(blok, lebar_isi, bawaan_ukuran, bawaan_font)
            if t is not None:
                cerita.append(KeepTogether(t))
                cerita.append(Spacer(1, 8))

    def halaman(canvas, _doc):
        canvas.saveState()
        gambar_kop(canvas, kop, A4[1])
        canvas.restoreState()

    pdf.build(cerita, onFirstPage=halaman, onLaterPages=halaman)
    buffer.seek(0)
    return buffer.getvalue()
