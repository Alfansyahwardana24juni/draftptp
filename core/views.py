import sqlite3, io, os, json
from django.shortcuts import render, redirect, get_object_or_404
from django.http import JsonResponse, FileResponse
from django.views.decorators.http import require_POST
from django.conf import settings
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from django.contrib.auth.models import User
from core.models import AktaPT, Pengurus, CatatanAkta, DraftUserStatus
from docxtpl import DocxTemplate
from datetime import datetime, date

TEMPLATE_PT_PERORANGAN = 'template_pt_perorangan.docx'


# ================= HELPER =================


def format_rupiah(n):
    if not n:
        return "0"
    return f"{int(n):,}".replace(",", ".")


def terbilang(n):
    n = int(n)
    bilang = ["", "satu", "dua", "tiga", "empat", "lima", "enam",
              "tujuh", "delapan", "sembilan", "sepuluh", "sebelas"]

    def _t(n):
        if n == 0: return ""
        elif n < 12: return bilang[n]
        elif n < 20: return _t(n - 10) + " belas"
        elif n < 100:
            sisa = (" " + _t(n % 10)) if n % 10 != 0 else ""
            return _t(n // 10) + " puluh" + sisa
        elif n < 200:
            sisa = (" " + _t(n - 100)) if n - 100 != 0 else ""
            return "seratus" + sisa
        elif n < 1000:
            sisa = (" " + _t(n % 100)) if n % 100 != 0 else ""
            return _t(n // 100) + " ratus" + sisa
        elif n < 2000:
            sisa = (" " + _t(n - 1000)) if n - 1000 != 0 else ""
            return "seribu" + sisa
        elif n < 1_000_000:
            sisa = (" " + _t(n % 1000)) if n % 1000 != 0 else ""
            return _t(n // 1000) + " ribu" + sisa
        elif n < 1_000_000_000:
            sisa = (" " + _t(n % 1_000_000)) if n % 1_000_000 != 0 else ""
            return _t(n // 1_000_000) + " juta" + sisa
        elif n < 1_000_000_000_000:
            sisa = (" " + _t(n % 1_000_000_000)) if n % 1_000_000_000 != 0 else ""
            return _t(n // 1_000_000_000) + " miliar" + sisa
        else:
            return "angka terlalu besar"

    return _t(n)


def format_tgl_indo(tgl):
    if not tgl:
        return ""
    try:
        dt = datetime.strptime(tgl, "%Y-%m-%d")
        bulan_nama = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
                      "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
        tgl_angka = dt.strftime("%d-%m-%Y")
        tgl_huruf = f"{terbilang(dt.day)} {bulan_nama[dt.month]} {terbilang(dt.year)}"
        return f"{tgl_angka} ({tgl_huruf})"
    except Exception:
        return tgl


def potong_kelebihan_panjang(obj):
    """Pangkas nilai CharField yang melebihi max_length.

    SQLite membiarkan nilai kepanjangan tersimpan, tapi Postgres menolaknya
    ("value too long for type character varying"). Dipangkas di sisi server
    supaya data tetap valid di kedua database.
    """
    for f in obj._meta.get_fields():
        max_len = getattr(f, 'max_length', None)
        if max_len and f.get_internal_type() == 'CharField':
            val = getattr(obj, f.name, None)
            if isinstance(val, str) and len(val) > max_len:
                setattr(obj, f.name, val[:max_len])
    return obj


def cek_kelengkapan(akta, pengurus_list):
    """Catatan kelengkapan draft PT Perorangan."""
    notes = []

    if (not akta.kota_pt or not akta.jalan_pt or not akta.rt_pt or not akta.rw_pt
            or not akta.desa_pt or not akta.kecamatan_pt or not akta.provinsi_pt):
        notes.append("Konfirmasi alamat lengkap PT")

    if not akta.kbli_json or akta.kbli_json == '[]':
        notes.append("Bidang usaha (KBLI) belum dipilih")

    if not pengurus_list:
        notes.append("Belum ada pengurus sama sekali")
        return notes

    for p in pengurus_list:
        nama = p.nama or 'Pengurus tanpa nama'
        # PT Perorangan tidak perlu cek tempat lahir
        if (not p.tgl_lahir or not p.nik or not p.pekerjaan
                or not p.kab or not p.jalan or not p.rt or not p.rw
                or not p.desa or not p.kecamatan or not p.provinsi):
            notes.append(f"{nama}: Konfirmasi kelengkapan identitas KTP")

    return notes


def build_pengurus_context(p, total_lembar, harga_saham=1_000_000):
    persen = int(p.saham_persen or 0)
    lembar = int((persen / 100) * total_lembar)
    nominal = lembar * harga_saham
    tgl_raw = p.tgl_lahir
    if tgl_raw:
        tgl_str = tgl_raw.strftime("%Y-%m-%d")
        tgl_fmt = format_tgl_indo(tgl_str)
    else:
        tgl_fmt = "**-**-**** (tanggal lahir belum diisi)"
    nik = p.nik.strip() if p.nik else "737*************"
    tpl = p.tpl or "[tempat lahir]"
    pekerjaan = p.pekerjaan or "[pekerjaan]"
    tipe_kab = p.tipe_kab or "Kabupaten"
    kab = p.kab or "[kota/kabupaten]"
    jalan = p.jalan or "[jalan]"
    rt = p.rt or "***"
    rw = p.rw or "***"
    desa = p.desa or "[desa/kelurahan]"
    kecamatan = p.kecamatan or "[kecamatan]"
    provinsi = p.provinsi or "[provinsi]"

    # Gabungkan gelar ke dalam nama
    nama_asli = (p.nama or "[nama]").upper().replace('TUAN ', '').replace('NYONYA ', '').replace('NONA ', '').strip()
    gelar_depan = p.gelar_depan.strip() if p.gelar_depan else ''
    gelar_belakang = p.gelar_belakang.strip() if p.gelar_belakang else ''
    gelar_singkat = p.gelar_singkat.strip() if p.gelar_singkat else ''

    # nama lengkap untuk bagian data pengurus (gelar panjang)
    if gelar_depan and gelar_belakang:
        nama_lengkap = f"{gelar_depan}. {nama_asli}, {gelar_belakang}"
    elif gelar_depan:
        nama_lengkap = f"{gelar_depan}. {nama_asli}"
    elif gelar_belakang:
        nama_lengkap = f"{nama_asli}, {gelar_belakang}"
    else:
        nama_lengkap = nama_asli

    # nama singkat untuk bagian penghadap tanda tangan
    singkat = gelar_singkat or gelar_belakang  # fallback ke gelar_belakang jika singkat kosong
    if gelar_depan and singkat:
        nama_singkat_otomatis = f"{gelar_depan}. {nama_asli}, {singkat}"
    elif gelar_depan:
        nama_singkat_otomatis = f"{gelar_depan}. {nama_asli}"
    elif singkat:
        nama_singkat_otomatis = f"{nama_asli}, {singkat}"
    else:
        nama_singkat_otomatis = nama_asli
        
    # Gunakan edit manual (kategori_yayasan) untuk PT PERORANGAN jika diisi
    if p.kategori_yayasan and getattr(p, 'akta', None) and getattr(p.akta, 'jenis_akta', '') == 'PT_PERORANGAN':
        nama_singkat = p.kategori_yayasan
    else:
        nama_singkat = nama_singkat_otomatis
    sebutan_val = p.sebutan or "Tuan"
    sbtn = "TN." if sebutan_val.upper() == "TUAN" else "NY." if sebutan_val.upper() == "NYONYA" else "NN." if sebutan_val.upper() == "NONA" else sebutan_val
    return {
        "sebutan": sebutan_val,
        "sbtn": sbtn,
        "nama": nama_lengkap,
        "pendiri": nama_lengkap,
        "nama_singkat": nama_singkat,
        "jabatan": p.jabatan or "[jabatan]",
        "persen": str(persen),
        "lembar": f"{lembar:,}".replace(",", "."),
        "terbilang_l": terbilang(lembar),
        "nominal": format_rupiah(nominal),
        "terbilang_nom": terbilang(nominal),
        "tpl": tpl,
        "tgl": tgl_fmt,
        "tgl_singkat": tgl_raw.strftime("%d-%m-%Y") if tgl_raw else "",
        "nik": nik,
        "npwp": p.npwp or '',
        "rt": rt,
        "desa": desa,
        "kecamatan": kecamatan,
        "tipe_kab": tipe_kab,
        "kab": kab,
        "provinsi": provinsi,
        "rw": rw, 
        "kerja": pekerjaan,
        "jalan": jalan,
        "kategori_yayasan": p.kategori_yayasan or '',
        "is_pendiri": p.is_pendiri,
        "alamat_lengkap": (
            f"{tipe_kab} {kab}, {jalan}, "
            f"Rukun Tetangga {rt}, Rukun Warga {rw}, "
            f"Kelurahan/Desa {desa}, Kecamatan {kecamatan}, "
            f"Pemegang Kartu Tanda Penduduk Provinsi {provinsi}, "
            f"{tipe_kab} {kab}"
        )
    }


def build_kbli_pt_perorangan(kodes, db_path):
    """
    PT NOT TANGSEL — Struktur flat (List Tunggal)
    Sama seperti build_kbli_palu namun biasanya PT membutuhkan field 
    yang lebih ringkas sesuai template PT NOT TANGSEL.
    """
    conn = sqlite3.connect(f'file:{db_path}?mode=ro', uri=True)
    cur = conn.cursor()
    list_kbli_final = []
    
    # Menghapus duplikat dan mengurutkan kode
    for kode in sorted(set(kodes)):
        if not kode: continue
        kode_str = str(kode).zfill(5)
        
        cur.execute("SELECT judul, deskripsi FROM tabel_kbli WHERE kode = ?", (kode_str,))
        row = cur.fetchone()
        
        if row:
            list_kbli_final.append({
                "kode": kode_str,
                "judul": row[0].title(),
                "deskripsi": row[1] if row[1] else ""
            })
            
    conn.close()
    return list_kbli_final


def simpan_ke_db(request, akta_id=None):
    nama_pt = request.POST.get('nama_pt', '').upper()
    kota_pt = request.POST.get('kota_pt', '').title()
    alamat_pt = request.POST.get('alamat_pt', '')
    jalan_pt = request.POST.get('jalan_pt', '').strip()
    rt_pt = request.POST.get('rt_pt', '').strip()
    rw_pt = request.POST.get('rw_pt', '').strip()
    desa_pt = request.POST.get('desa_pt', '').strip()
    kecamatan_pt = request.POST.get('kecamatan_pt', '').strip()
    tipe_kab_pt = request.POST.get('tipe_kab_pt', 'Kota').strip() or 'Kota'
    provinsi_pt = request.POST.get('provinsi_pt', '').strip()
    pos_pt = request.POST.get('pos_pt', '').strip()
    modal_input = int(request.POST.get('modal_dasar_input', 0) or 0)
    persen_disetor = int(request.POST.get('persen_modal_disetor', 100) or 100)
    harga_saham_input = int(request.POST.get('harga_saham_input', 1_000_000) or 1_000_000)
    total_lembar = int(modal_input / harga_saham_input) if harga_saham_input > 0 else 0
    kbli_kodes = [k for k in request.POST.getlist('kbli_kode[]') if k]
    kbli_json = json.dumps(kbli_kodes)
    kantor_kepaniteraan = request.POST.get('kantor_kepaniteraan', '').strip()
    kota_kepaniteraan = request.POST.get('kota_kepaniteraan', '').strip()
    tipe_kepaniteraan = request.POST.get('tipe_kepaniteraan', 'Kota').strip()
    
    # Default jika tidak diisi
    if not kantor_kepaniteraan:
        kantor_kepaniteraan = 'Kantor Kepaniteraan Pengadilan Negeri Makassar'
        kota_kepaniteraan = 'Makassar'
        tipe_kepaniteraan = 'Kota'

    if akta_id:
        akta = get_object_or_404(AktaPT, pk=akta_id)
        akta.nama_pt = nama_pt
        akta.kota_pt = kota_pt
        akta.alamat_pt = alamat_pt
        akta.jalan_pt = jalan_pt
        akta.rt_pt = rt_pt
        akta.rw_pt = rw_pt
        akta.desa_pt = desa_pt
        akta.kecamatan_pt = kecamatan_pt
        akta.tipe_kab_pt = tipe_kab_pt
        akta.provinsi_pt = provinsi_pt
        akta.pos_pt = pos_pt
        akta.modal_dasar = modal_input
        akta.persen_modal_disetor = persen_disetor
        akta.total_lembar = total_lembar
        akta.kbli_json = kbli_json
        akta.kantor_kepaniteraan = kantor_kepaniteraan
        akta.kota_kepaniteraan = kota_kepaniteraan
        akta.tipe_kepaniteraan = tipe_kepaniteraan
        if not getattr(akta, 'user', None) and request.user.is_authenticated:
            akta.user = request.user
        potong_kelebihan_panjang(akta)
        akta.save()
        akta.pengurus_set.all().delete()
    else:
        akta = AktaPT(
            nama_pt=nama_pt, kota_pt=kota_pt, alamat_pt=alamat_pt,
            jalan_pt=jalan_pt, rt_pt=rt_pt, rw_pt=rw_pt,
            desa_pt=desa_pt, kecamatan_pt=kecamatan_pt, tipe_kab_pt=tipe_kab_pt,
            provinsi_pt=provinsi_pt, pos_pt=pos_pt,
            modal_dasar=modal_input, persen_modal_disetor=persen_disetor, total_lembar=total_lembar,
            kbli_json=kbli_json, jenis_akta='PT_PERORANGAN',
            kantor_kepaniteraan=kantor_kepaniteraan,
            kota_kepaniteraan=kota_kepaniteraan,
            tipe_kepaniteraan=tipe_kepaniteraan,
            user=request.user if request.user.is_authenticated else None
        )
        potong_kelebihan_panjang(akta)
        akta.save()

    names = request.POST.getlist('nama[]')
    sebutan_l = request.POST.getlist('sebutan[]')
    jabatan_l = request.POST.getlist('jabatan[]')
    persen_l = request.POST.getlist('saham_persen[]')
    kategori_yayasan_l = request.POST.getlist('kategori_yayasan[]')
    is_pendiri_l = request.POST.getlist('is_pendiri[]')
    gelar_depan_l = request.POST.getlist('gelar_depan[]')
    gelar_belakang_l = request.POST.getlist('gelar_belakang[]')
    gelar_singkat_l = request.POST.getlist('gelar_singkat[]')
    tpl_l = request.POST.getlist('tpl[]')
    tgl_l = request.POST.getlist('tgl[]')
    nik_l = request.POST.getlist('nik[]')
    npwp_l = request.POST.getlist('npwp[]')
    kerja_l = request.POST.getlist('kerja[]')
    tipe_kab_l = request.POST.getlist('tipe_kab[]')
    kab_l = request.POST.getlist('kab[]')
    jalan_l = request.POST.getlist('jalan[]')
    rt_l = request.POST.getlist('rt[]')
    rw_l = request.POST.getlist('rw[]')
    desa_l = request.POST.getlist('desa[]')
    kec_l = request.POST.getlist('kec[]')
    prov_l = request.POST.getlist('prov[]')

    for i in range(len(names)):
        tgl_val = tgl_l[i] if tgl_l[i] else None
        # Ubah jabatan menjadi Title Case (Managing Partner, Partner)
        jabatan_val = jabatan_l[i].title() if jabatan_l[i] else ''
        pengurus = Pengurus(
            akta=akta,
            sebutan=sebutan_l[i], nama=names[i].upper(), jabatan=jabatan_val,
            saham_persen=int(persen_l[i] or 0),
            kategori_yayasan=kategori_yayasan_l[i] if i < len(kategori_yayasan_l) else '',
            is_pendiri=True if (i < len(is_pendiri_l) and is_pendiri_l[i] == '1') else False,
            gelar_depan=gelar_depan_l[i] if i < len(gelar_depan_l) else '',
            gelar_belakang=gelar_belakang_l[i] if i < len(gelar_belakang_l) else '',
            gelar_singkat=gelar_singkat_l[i] if i < len(gelar_singkat_l) else '',
            tpl=tpl_l[i], tgl_lahir=tgl_val,
            nik=nik_l[i], npwp=npwp_l[i] if i < len(npwp_l) else '', pekerjaan=kerja_l[i], tipe_kab=tipe_kab_l[i],
            kab=kab_l[i], jalan=jalan_l[i], rt=rt_l[i], rw=rw_l[i],
            desa=desa_l[i], kecamatan=kec_l[i], provinsi=prov_l[i]
        )
        potong_kelebihan_panjang(pengurus)
        pengurus.save()

    # Track DraftUserStatus points
    if request.user.is_authenticated:
        pengurus_list = list(akta.pengurus_set.all())
        notes = cek_kelengkapan(akta, pengurus_list)
        is_lengkap = (len(notes) == 0)
        is_revisi = (akta_id is not None)  # True jika edit, False jika draf baru
        DraftUserStatus.objects.update_or_create(
            akta=akta, user=request.user,
            defaults={'is_lengkap': is_lengkap, 'is_revisi': is_revisi, 'date_recorded': date.today()}
        )

    return akta.id


# ================= VIEWS =================


def login_view(request):
    error = None
    username_val = ""
    if request.method == 'POST':
        username_val = request.POST.get('username', '')
        password = request.POST.get('password', '')
        user = authenticate(request, username=username_val, password=password)
        if user is not None:
            login(request, user)
            return redirect('draft_pt_perorangan')
        else:
            # Jika user tidak ditemukan, kosongkan input username (refresh form)
            if not User.objects.filter(username=username_val).exists():
                username_val = ""
            error = 'Username atau password salah'
    return render(request, 'login.html', {'error': error, 'username': username_val})


def logout_view(request):
    logout(request)
    return redirect('login')


@login_required(login_url='/login/')
def draft_pt_perorangan(request):
    return render(request, 'draft_pt_perorangan.html')


@login_required(login_url='/login/')
def riwayat(request):
    akta_list = AktaPT.objects.filter(jenis_akta='PT_PERORANGAN').order_by('-updated_at')
    result = []
    for akta in akta_list:
        pengurus_list = list(akta.pengurus_set.all())
        catatan_list = list(akta.catatan_set.filter(selesai=False).order_by('created_at'))
        kbli_raw = json.loads(akta.kbli_json or '[]')
        kbli_sorted = sorted(set(kbli_raw))
        kbli_rows = [kbli_sorted[i:i+5] for i in range(0, len(kbli_sorted), 5)]
        notes = cek_kelengkapan(akta, pengurus_list)
        result.append({
            'akta': akta,
            'pengurus': [p.to_dict() for p in pengurus_list],
            'notes': notes,
            'lengkap': len(notes) == 0 and len(catatan_list) == 0,
            'catatan': catatan_list,
            'kbli_rows': kbli_rows,
            'modal_dasar_fmt': f"{int(akta.modal_dasar):,}".replace(",", ".") if akta.modal_dasar else '0',
        })
    return render(request, 'riwayat.html', {'data': result})


@login_required(login_url='/login/')
def edit(request, akta_id):
    akta = get_object_or_404(AktaPT, pk=akta_id, jenis_akta='PT_PERORANGAN')
    pengurus_list = [p.to_dict() for p in akta.pengurus_set.all()]
    akta_dict = {
        'id': akta.id, 'nama_pt': akta.nama_pt, 'kota_pt': akta.kota_pt,
        'alamat_pt': akta.alamat_pt, 'modal_dasar': akta.modal_dasar,
        'persen_modal_disetor': akta.persen_modal_disetor,
        'total_lembar': akta.total_lembar, 'kbli_json': akta.kbli_json,
        'kantor_kepaniteraan': akta.kantor_kepaniteraan or '',
        'kota_kepaniteraan': akta.kota_kepaniteraan or '',
        'tipe_kepaniteraan': akta.tipe_kepaniteraan or '',
        'jalan_pt': akta.jalan_pt or '',
        'rt_pt': akta.rt_pt or '',
        'rw_pt': akta.rw_pt or '',
        'desa_pt': akta.desa_pt or '',
        'kecamatan_pt': akta.kecamatan_pt or '',
        'tipe_kab_pt': akta.tipe_kab_pt or 'Kota',
        'provinsi_pt': akta.provinsi_pt or '',
        'pos_pt': akta.pos_pt or '',
    }
    return render(request, 'draft_pt_perorangan.html', {
        'akta': akta,
        'akta_json': json.dumps(akta_dict),
        'pengurus_json': json.dumps(pengurus_list),
        'edit_mode': True,
        'auto_download': request.GET.get('download') == '1',
        'akta_id': akta_id,
    })


@login_required(login_url='/login/')
def cari_kbli(request):
    q = request.GET.get('q', '').strip()
    if len(q) < 2:
        return JsonResponse([], safe=False)
    db_path = os.path.join(settings.BASE_DIR, 'kbli.db')
    # mode=ro: kbli.db hanya dibaca, dan di serverless filesystem-nya read-only
    conn = sqlite3.connect(f'file:{db_path}?mode=ro', uri=True)
    cursor = conn.cursor()
    cursor.execute(
        "SELECT kode, judul FROM tabel_kbli WHERE kode LIKE ? OR judul LIKE ? LIMIT 10",
        (f'%{q}%', f'%{q}%')
    )
    data = [{"kode": r[0], "judul": r[1]} for r in cursor.fetchall()]
    conn.close()
    return JsonResponse(data, safe=False)


@login_required(login_url='/login/')
@require_POST
def proses(request):
    akta_id_edit = request.POST.get('akta_id_edit')
    akta_id = simpan_ke_db(request, int(akta_id_edit) if akta_id_edit else None)
    return redirect(f'/edit/{akta_id}/?download=1')


@login_required(login_url='/login/')
@require_POST
def proses_simpan(request):
    akta_id_edit = request.POST.get('akta_id_edit')
    simpan_ke_db(request, int(akta_id_edit) if akta_id_edit else None)
    return redirect('riwayat')


@login_required(login_url='/login/')
def download(request, akta_id):
    akta = get_object_or_404(AktaPT, pk=akta_id, jenis_akta='PT_PERORANGAN')
    pengurus_rows = list(akta.pengurus_set.all())

    modal_input = akta.modal_dasar
    total_lembar = akta.total_lembar
    persen_disetor = akta.persen_modal_disetor or 100
    harga_saham = (akta.modal_dasar // akta.total_lembar) if akta.total_lembar > 0 else 1_000_000
    lembar_disetor = int(total_lembar * persen_disetor / 100)
    modal_disetor = int(modal_input * persen_disetor / 100)

    pengurus_data = [build_pengurus_context(p, lembar_disetor, harga_saham) for p in pengurus_rows]

    kodes = json.loads(akta.kbli_json or '[]')
    db_path = os.path.join(settings.BASE_DIR, 'kbli.db')
    data_kbli = build_kbli_pt_perorangan(kodes, db_path)
    if not data_kbli:
        data_kbli = [{"kode": "undefined", "judul": "UNDEFINED", "deskripsi": "undefined"}]

    context = {
        "NAMA_PT": akta.nama_pt,
        "MODAL_DISETOR": format_rupiah(modal_disetor),
        "data_kbli": data_kbli,
        "pengurus": pengurus_data,
    }

    if pengurus_data:
        p0 = pengurus_data[0]
        # Format tgl lahir: 01/01/1985 (dengan nol di depan)
        tgl_indo = ''
        if p0['tgl_singkat']:
            try:
                d = datetime.strptime(p0['tgl_singkat'], "%d-%m-%Y")
                tgl_indo = d.strftime("%d/%m/%Y")
            except Exception:
                tgl_indo = p0['tgl_singkat'].replace('-', '/')

        alamat_khusus = (
            f"{p0['jalan']}, RT {p0['rt']}, RW {p0['rw']}, {p0['desa']}, "
            f"{p0['kecamatan']}, {p0['tipe_kab']} {p0['kab']}, {p0['provinsi']}"
        )

        context.update({
            "nama": p0['nama'],
            "nama_singkat": p0['nama_singkat'],
            "tgl_lahir": tgl_indo,
            "ALAMAT_PENGURUS": alamat_khusus,
            "nik": p0['nik'],
            "npwp": p0['npwp'],
        })

    context.update({
        "JLN_PT": akta.jalan_pt or '[Jalan]',
        "RT_PT": akta.rt_pt or '***',
        "RW_PT": akta.rw_pt or '***',
        "DESA_PT": akta.desa_pt or '[Desa/Kelurahan]',
        "KECAMATAN_PT": akta.kecamatan_pt or '[Kecamatan]',
        "KOTA_PT": f"{akta.tipe_kab_pt or 'Kota'} {akta.kota_pt}",
        "PROVINSI_PT": akta.provinsi_pt or '[Provinsi]',
    })

    template_path = os.path.join(settings.BASE_DIR, 'template_word', TEMPLATE_PT_PERORANGAN)
    doc = DocxTemplate(template_path)
    doc.render(context)
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    # Sanitize filename — ganti karakter yang bermasalah
    safe_name = akta.nama_pt.replace('"', '').replace("'", "").replace('/', '-').replace('\\', '-')
    filename = f"DRAFT_{safe_name}.docx"

    response = FileResponse(
        buffer,
        content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    response['Content-Length'] = buffer.getbuffer().nbytes
    return response


@login_required
@require_POST
def get_kbli_recommendation(request):
    import os
    import json
    import csv
    import google.generativeai as genai
    from django.http import JsonResponse
    from django.conf import settings

    deskripsi = request.POST.get('deskripsi', '').strip()
    if not deskripsi:
        return JsonResponse({'error': 'Deskripsi usaha tidak boleh kosong'}, status=400)

    api_key = os.environ.get('GEMINI_API_KEY')
    if not api_key:
        return JsonResponse({'error': 'API Key belum dikonfigurasi di server'}, status=500)

    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-2.5-flash')
        
        prompt = '''
        Saya akan memberikan deskripsi sebuah kegiatan usaha. Tolong berikan 3 rekomendasi kode KBLI 2020 (5 digit) yang paling cocok untuk usaha tersebut, beserta nama/judul KBLI-nya.
        Format output harus murni berupa array of JSON objects, tanpa markdown tambahan.
        Contoh format:
        [
            {"kode": "56101", "judul": "KEDAI MAKANAN"},
            {"kode": "56303", "judul": "RUMAH MINUM/KAFE"}
        ]
        
        Deskripsi Usaha: ''' + deskripsi + '''
        '''
        
        response = model.generate_content(prompt)
        text = response.text.strip()
        
        if text.startswith('```json'):
            text = text[7:]
        if text.startswith('```'):
            text = text[3:]
        if text.endswith('```'):
            text = text[:-3]
        text = text.strip()
        
        data_2020 = json.loads(text)

        # ---------------- HYBRID DATABASE MAPPING (KBLI 2020 -> 2025) ----------------
        csv_path = os.path.join(settings.BASE_DIR, 'scripts', 'konversi_kbli_117_231_fix.csv')
        
        mapping_kbli = {}
        if os.path.exists(csv_path):
            with open(csv_path, mode='r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                next(reader, None) # Hanya 1 header
                
                for row in reader:
                    if len(row) >= 4:
                        kode_2020 = row[0].strip()
                        kode_2025 = row[2].strip()
                        judul_2025 = row[3].strip()
                        if kode_2020:
                            if kode_2020 not in mapping_kbli:
                                mapping_kbli[kode_2020] = []
                            mapping_kbli[kode_2020].append({
                                'kode': kode_2025,
                                'judul': judul_2025
                            })
        
        mapped_data = []
        for item in data_2020:
            kode_ai = item.get('kode', '').strip()
            if kode_ai in mapping_kbli:
                for mapped_item in mapping_kbli[kode_ai]:
                    if mapped_item not in mapped_data:
                        mapped_data.append(mapped_item)
            else:
                if item not in mapped_data:
                    mapped_data.append(item)
                
        return JsonResponse({'success': True, 'data': mapped_data})
        
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)


@login_required(login_url='/login/')
@require_POST
def catatan_tambah(request):
    akta_id = request.POST.get('akta_id')
    isi = request.POST.get('isi', '').strip()
    if isi and akta_id:
        akta = get_object_or_404(AktaPT, pk=akta_id)
        CatatanAkta.objects.create(akta=akta, isi=isi)
    return redirect('riwayat')


@login_required(login_url='/login/')
@require_POST
def catatan_selesai(request, catatan_id):
    CatatanAkta.objects.filter(pk=catatan_id).delete()
    return JsonResponse({'ok': True})


@login_required(login_url='/login/')
@require_POST
def hapus(request, akta_id):
    akta = get_object_or_404(AktaPT, pk=akta_id, jenis_akta='PT_PERORANGAN')
    akta.delete()
    return redirect('riwayat')
