"""
Ubah dokumen .docx hasil render menjadi PDF.

Konverter ini membaca dokumen yang BARU SAJA dibuat docxtpl, lalu menyusun
ulang isinya (paragraf dan tabel) ke PDF. Jadi PDF selalu mengikuti template
Word — kalau templatenya diubah, PDF ikut berubah tanpa menyentuh kode ini.

Catatan jujur: ini penyusunan ulang, bukan konversi piksel-per-piksel.
Teks, urutan, struktur tabel, dan perataan ikut; detail halus seperti spasi
antarbaris persis atau bingkai khas Word bisa sedikit berbeda. Konversi yang
benar-benar identik membutuhkan LibreOffice atau Microsoft Word, dan keduanya
tidak tersedia di lingkungan serverless.
"""
import io

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.table import Table
from docx.text.paragraph import Paragraph
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import KeepTogether, Paragraph as PdfParagraph
from reportlab.platypus import SimpleDocTemplate, Spacer, Table as PdfTable, TableStyle

RATA = {
    WD_ALIGN_PARAGRAPH.CENTER: TA_CENTER,
    WD_ALIGN_PARAGRAPH.RIGHT: TA_RIGHT,
    WD_ALIGN_PARAGRAPH.JUSTIFY: TA_JUSTIFY,
}

GARIS = colors.HexColor('#000000')


def _escape(teks):
    return (teks or '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def _inline(paragraf):
    """Susun teks paragraf sambil mempertahankan tebal dan miring."""
    bagian = []
    for run in paragraf.runs:
        potongan = _escape(run.text)
        if not potongan:
            continue
        if run.bold:
            potongan = f'<b>{potongan}</b>'
        if run.italic:
            potongan = f'<i>{potongan}</i>'
        bagian.append(potongan)
    return ''.join(bagian) or _escape(paragraf.text)


def _gaya(nama, ukuran, rata=TA_LEFT, tebal=False, spasi_bawah=4):
    return ParagraphStyle(
        nama,
        fontName='Times-Bold' if tebal else 'Times-Roman',
        fontSize=ukuran,
        leading=ukuran * 1.35,
        alignment=rata,
        spaceAfter=spasi_bawah,
    )


def _gaya_turunan(paragraf, ambil):
    """Telusuri gaya paragraf dan induknya sampai menemukan nilai.

    Word menyimpan format judul di gaya (mis. 'Heading 1'), bukan di run,
    jadi membaca run saja membuat judul kehilangan ukuran dan ketebalannya.
    """
    gaya = paragraf.style
    while gaya is not None:
        nilai = ambil(gaya)
        if nilai is not None:
            return nilai
        gaya = gaya.base_style
    return None


def _ukuran_font(paragraf, bawaan=9):
    for run in paragraf.runs:
        if run.font.size:
            return run.font.size.pt
    ukuran = _gaya_turunan(paragraf, lambda g: g.font.size)
    return ukuran.pt if ukuran else bawaan


def _tebal(paragraf):
    if any(r.bold for r in paragraf.runs if r.text):
        return True
    return bool(_gaya_turunan(paragraf, lambda g: g.font.bold))


def _rata(paragraf):
    if paragraf.alignment is not None:
        return RATA.get(paragraf.alignment, TA_LEFT)
    align = _gaya_turunan(paragraf, lambda g: g.paragraph_format.alignment)
    if align is not None:
        return RATA.get(align, TA_LEFT)
    # Judul tanpa perataan eksplisit di Word umumnya tampil di tengah.
    return TA_CENTER if paragraf.style.name.startswith('Heading') else TA_LEFT


def _isi_blok(induk):
    """Hasilkan paragraf dan tabel sesuai urutan aslinya di dokumen."""
    body = induk.element.body
    for anak in body.iterchildren():
        if anak.tag.endswith('}p'):
            yield Paragraph(anak, induk)
        elif anak.tag.endswith('}tbl'):
            yield Table(anak, induk)


def _sel_teks(sel, gaya):
    bagian = [_inline(p) for p in sel.paragraphs]
    teks = '<br/>'.join([b for b in bagian if b]) or ''
    return PdfParagraph(teks, gaya)


def _bangun_tabel(tabel, lebar_total):
    kolom = len(tabel.columns)
    if not kolom:
        return None

    gaya_sel = _gaya('sel', 9, spasi_bawah=0)
    gaya_kepala = _gaya('kepala', 9, rata=TA_CENTER, tebal=True, spasi_bawah=0)

    data = []
    perintah = [
        ('GRID', (0, 0), (-1, -1), 0.5, GARIS),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]

    for i, baris in enumerate(tabel.rows):
        sel_baris = list(baris.cells)
        # Sel yang digabung muncul berulang di python-docx; deteksi lewat
        # elemen tc yang sama supaya bisa dijadikan SPAN di PDF.
        terlihat = []
        mulai = 0
        for j, sel in enumerate(sel_baris):
            if j > 0 and sel._tc is sel_baris[j - 1]._tc:
                continue
            if terlihat:
                akhir = j - 1
                if akhir > mulai:
                    perintah.append(('SPAN', (mulai, i), (akhir, i)))
            terlihat.append(sel)
            mulai = j
        if mulai < len(sel_baris) - 1:
            perintah.append(('SPAN', (mulai, i), (len(sel_baris) - 1, i)))

        gaya = gaya_kepala if (i == 0 and len(terlihat) == 1) else gaya_sel
        isi = [_sel_teks(s, gaya) for s in terlihat]
        isi += [''] * (kolom - len(isi))
        data.append(isi)

        if i == 0 and len(terlihat) == 1:
            perintah.append(('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#e8e8e8')))

    # Lebar kolom: ikuti lebar asli dari Word kalau ada, kalau tidak dibagi rata.
    lebar = []
    try:
        for kol in tabel.columns:
            lebar.append(kol.width.pt if kol.width else None)
    except Exception:
        lebar = [None] * kolom
    if any(w is None for w in lebar) or not sum(w for w in lebar if w):
        lebar = [lebar_total / kolom] * kolom
    else:
        skala = lebar_total / sum(lebar)
        lebar = [w * skala for w in lebar]

    t = PdfTable(data, colWidths=lebar, repeatRows=0)
    t.setStyle(TableStyle(perintah))
    return t


def docx_ke_pdf(docx_bytes):
    """Terima byte .docx, balas byte PDF."""
    doc = Document(io.BytesIO(docx_bytes))
    sec = doc.sections[0]

    margin_kiri = sec.left_margin.pt if sec.left_margin else 20 * mm
    margin_kanan = sec.right_margin.pt if sec.right_margin else 20 * mm
    margin_atas = sec.top_margin.pt if sec.top_margin else 18 * mm
    margin_bawah = sec.bottom_margin.pt if sec.bottom_margin else 18 * mm

    buffer = io.BytesIO()
    pdf = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=margin_kiri,
        rightMargin=margin_kanan,
        topMargin=margin_atas,
        bottomMargin=margin_bawah,
        title=doc.paragraphs[0].text if doc.paragraphs else 'Dokumen',
    )
    lebar_isi = A4[0] - margin_kiri - margin_kanan

    cerita = []
    for blok in _isi_blok(doc):
        if isinstance(blok, Paragraph):
            teks = _inline(blok)
            if not teks.strip():
                cerita.append(Spacer(1, 4))
                continue
            cerita.append(PdfParagraph(teks, _gaya(
                'p', _ukuran_font(blok), _rata(blok), _tebal(blok), 6)))
        else:
            t = _bangun_tabel(blok, lebar_isi)
            if t is not None:
                cerita.append(KeepTogether(t))
                cerita.append(Spacer(1, 8))

    pdf.build(cerita)
    buffer.seek(0)
    return buffer.getvalue()
