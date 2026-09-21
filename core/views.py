"""
VELLOSCRIPT — Draft PT Perorangan.

Aplikasi ini TIDAK memakai database. Draft disimpan di browser pengguna
(localStorage); server hanya bertugas dua hal:

  1. mencari kode KBLI dari kbli.db (dibaca saja, tidak pernah ditulis), dan
  2. merender dokumen .docx dari template Word memakai docxtpl.

Keduanya dijalankan di server karena penanda di dalam file Word terpecah
ke beberapa bagian XML — docxtpl menanganinya dengan benar, sementara
implementasi di sisi browser rawan menghasilkan dokumen rusak.
"""
import io
import os
import re
import sqlite3
from datetime import datetime

from django.conf import settings
from django.http import FileResponse, JsonResponse
from django.shortcuts import render
from django.views.decorators.http import require_POST
from docxtpl import DocxTemplate

TEMPLATE_PT_PERORANGAN = 'template_pt_perorangan.docx'


# ================= HELPER =================

def kbli_db_path():
    return os.path.join(settings.BASE_DIR, 'kbli.db')


def buka_kbli():
    """kbli.db hanya dibaca; mode=ro supaya aman di filesystem read-only."""
    return sqlite3.connect(f'file:{kbli_db_path()}?mode=ro', uri=True)


def format_rupiah(n):
    """Angka saja dengan pemisah titik — template Word sudah memuat "Rp." sendiri."""
    try:
        n = int(n)
    except (TypeError, ValueError):
        n = 0
    if not n:
        return '0'
    return f'{n:,}'.replace(',', '.')


def ambil(request, nama, bawaan=''):
    return (request.POST.get(nama) or bawaan).strip()


def rangkai_nama(nama, gelar_depan, gelar_belakang):
    """Susun nama lengkap dengan gelar, tanda baca dirangkai otomatis."""
    inti = (nama or '[nama]').upper()
    for sebutan in ('TUAN ', 'NYONYA ', 'NONA '):
        inti = inti.replace(sebutan, '')
    inti = inti.strip()

    # Pengguna sering sudah mengetik titik ("Drs."), jangan ditambahi lagi.
    if gelar_depan and not gelar_depan.endswith('.'):
        gelar_depan += '.'

    if gelar_depan and gelar_belakang:
        return f'{gelar_depan} {inti}, {gelar_belakang}'
    if gelar_depan:
        return f'{gelar_depan} {inti}'
    if gelar_belakang:
        return f'{inti}, {gelar_belakang}'
    return inti


def format_tanggal(nilai):
    """'1985-01-01' -> '01/01/1985'. Kosong atau tidak terbaca -> string kosong."""
    if not nilai:
        return ''
    try:
        return datetime.strptime(nilai, '%Y-%m-%d').strftime('%d/%m/%Y')
    except ValueError:
        return nilai


def build_kbli(kodes):
    """Ambil judul & deskripsi tiap kode KBLI, tanpa duplikat, terurut."""
    hasil = []
    conn = buka_kbli()
    cur = conn.cursor()
    for kode in sorted({k for k in kodes if k}):
        kode_str = str(kode).zfill(5)
        cur.execute('SELECT judul, deskripsi FROM tabel_kbli WHERE kode = ?', (kode_str,))
        row = cur.fetchone()
        if row:
            hasil.append({
                'kode': kode_str,
                'judul': (row[0] or '').title(),
                'deskripsi': row[1] or '',
            })
    conn.close()
    return hasil


def build_context(request):
    """Rakit seluruh variabel yang dibutuhkan template Word dari data form."""
    nama_pt = ambil(request, 'nama_pt').upper()

    try:
        modal = int(ambil(request, 'modal_dasar', '0') or 0)
    except ValueError:
        modal = 0
    try:
        persen = int(ambil(request, 'persen_disetor', '100') or 100)
    except ValueError:
        persen = 100
    persen = max(0, min(100, persen))
    modal_disetor = int(modal * persen / 100)

    data_kbli = build_kbli(request.POST.getlist('kbli_kode[]'))
    if not data_kbli:
        data_kbli = [{'kode': 'undefined', 'judul': 'UNDEFINED', 'deskripsi': 'undefined'}]

    # --- pemilik ---
    gelar_depan = ambil(request, 'gelar_depan')
    gelar_belakang = ambil(request, 'gelar_belakang')
    gelar_singkat = ambil(request, 'gelar_singkat')
    nama_lengkap = rangkai_nama(ambil(request, 'nama'), gelar_depan, gelar_belakang)

    # Untuk bagian penghadap dipakai gelar singkat; bisa ditimpa manual.
    singkat = gelar_singkat or gelar_belakang
    nama_singkat = rangkai_nama(ambil(request, 'nama'), gelar_depan, singkat)
    manual = ambil(request, 'nama_penghadap')
    if manual:
        nama_singkat = manual

    tipe_kab = ambil(request, 'tipe_kab', 'Kabupaten')
    kab = ambil(request, 'kab') or '[kota/kabupaten]'
    alamat_pengurus = ', '.join([
        ambil(request, 'jalan') or '[jalan]',
        'RT ' + (ambil(request, 'rt') or '***'),
        'RW ' + (ambil(request, 'rw') or '***'),
        ambil(request, 'desa') or '[desa/kelurahan]',
        ambil(request, 'kecamatan') or '[kecamatan]',
        f'{tipe_kab} {kab}',
        ambil(request, 'provinsi') or '[provinsi]',
    ])

    return {
        'NAMA_PT': nama_pt or '[NAMA PT]',
        'MODAL_DISETOR': format_rupiah(modal_disetor),
        'data_kbli': data_kbli,

        'JLN_PT': ambil(request, 'jalan_pt') or '[Jalan]',
        'RT_PT': ambil(request, 'rt_pt') or '***',
        'RW_PT': ambil(request, 'rw_pt') or '***',
        'DESA_PT': ambil(request, 'desa_pt') or '[Desa/Kelurahan]',
        'KECAMATAN_PT': ambil(request, 'kecamatan_pt') or '[Kecamatan]',
        'KOTA_PT': f"{ambil(request, 'tipe_kab_pt', 'Kota')} {ambil(request, 'kota_pt')}".strip(),
        'PROVINSI_PT': ambil(request, 'provinsi_pt') or '[Provinsi]',

        'nama': nama_lengkap,
        'nama_singkat': nama_singkat,
        'tgl_lahir': format_tanggal(ambil(request, 'tgl_lahir')),
        'nik': ambil(request, 'nik') or '737*************',
        'npwp': ambil(request, 'npwp'),
        'ALAMAT_PENGURUS': alamat_pengurus,
    }


def nama_berkas(nama_pt, ekstensi):
    bersih = re.sub(r'[\\/:*?"<>|]', '-', nama_pt or 'DRAFT').strip() or 'DRAFT'
    return f'DRAFT_{bersih}.{ekstensi}'


# ================= VIEWS =================

def form(request):
    """Halaman formulir. Draft dimuat & disimpan oleh JavaScript di browser."""
    return render(request, 'form.html')


def riwayat(request):
    """Daftar draft — isinya dirender di browser dari localStorage."""
    return render(request, 'riwayat.html')


def cari_kbli(request):
    q = (request.GET.get('q') or '').strip()
    if len(q) < 2:
        return JsonResponse([], safe=False)
    conn = buka_kbli()
    cur = conn.cursor()
    cur.execute(
        'SELECT kode, judul FROM tabel_kbli WHERE kode LIKE ? OR judul LIKE ? LIMIT 10',
        (f'%{q}%', f'%{q}%'),
    )
    data = [{'kode': r[0], 'judul': r[1]} for r in cur.fetchall()]
    conn.close()
    return JsonResponse(data, safe=False)


def render_docx(request):
    """Render template Word dengan data form, balas byte .docx."""
    context = build_context(request)
    template_path = os.path.join(settings.BASE_DIR, 'template_word', TEMPLATE_PT_PERORANGAN)
    doc = DocxTemplate(template_path)
    # autoescape wajib: tanpa ini karakter & < > dari isian pengguna
    # disuntikkan mentah ke XML dan merusak struktur dokumen.
    doc.render(context, autoescape=True)

    buffer = io.BytesIO()
    doc.save(buffer)
    return context, buffer.getvalue()


def kirim_berkas(isi, nama, tipe):
    response = FileResponse(io.BytesIO(isi), content_type=tipe)
    response['Content-Disposition'] = f'attachment; filename="{nama}"'
    response['Content-Length'] = len(isi)
    return response


@require_POST
def generate(request):
    """Terima data form, balas dokumen .docx. Tidak ada yang disimpan di server."""
    context, isi = render_docx(request)
    return kirim_berkas(
        isi,
        nama_berkas(context['NAMA_PT'], 'docx'),
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    )


@require_POST
def generate_pdf(request):
    """Sama seperti generate(), tapi dokumennya diubah dulu menjadi PDF."""
    from .pdf import docx_ke_pdf

    context, isi = render_docx(request)
    return kirim_berkas(isi=docx_ke_pdf(isi),
                        nama=nama_berkas(context['NAMA_PT'], 'pdf'),
                        tipe='application/pdf')
