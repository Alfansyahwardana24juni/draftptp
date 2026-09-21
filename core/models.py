from django.db import models
from django.contrib.auth.models import User


class AktaPT(models.Model):
    JENIS_CHOICES = [
        ('PT_PERORANGAN', 'PT PERORANGAN'),
    ]
    jenis_akta = models.CharField(max_length=50, choices=JENIS_CHOICES, default='PT_PERORANGAN')
    nama_pt = models.CharField(max_length=255)
    kota_pt = models.CharField(max_length=100, blank=True)
    alamat_pt = models.TextField(blank=True)
    jalan_pt = models.TextField(blank=True)
    rt_pt = models.CharField(max_length=10, blank=True)
    rw_pt = models.CharField(max_length=10, blank=True)
    desa_pt = models.CharField(max_length=100, blank=True)
    kecamatan_pt = models.CharField(max_length=100, blank=True)
    tipe_kab_pt = models.CharField(max_length=20, blank=True, default='Kota')
    provinsi_pt = models.CharField(max_length=100, blank=True)
    pos_pt = models.CharField(max_length=10, blank=True)
    modal_dasar = models.BigIntegerField(default=0)
    persen_modal_disetor = models.IntegerField(default=100)
    total_lembar = models.IntegerField(default=0)
    kbli_json = models.TextField(default='[]')
    kantor_kepaniteraan = models.CharField(max_length=255, blank=True, default='')
    kota_kepaniteraan = models.CharField(max_length=100, blank=True, default='')
    tipe_kepaniteraan = models.CharField(max_length=20, blank=True, default='Kota')
    provinsi_yayasan = models.CharField(max_length=100, blank=True, default='')
    desa_cv = models.CharField(max_length=100, blank=True, default='')
    kecamatan_cv = models.CharField(max_length=100, blank=True, default='')
    email_pt = models.CharField(max_length=100, blank=True, default='')
    password_email_pt = models.CharField(max_length=100, blank=True, default='')
    telp_pt = models.CharField(max_length=50, blank=True, default='')
    status = models.CharField(max_length=20, default='draft')
    user = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='akta_set')
    is_surat_kuasa_generated = models.BooleanField(default=False)
    surat_kuasa_generated_at = models.DateTimeField(null=True, blank=True)
    surat_kuasa_user = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='surat_kuasa_set')
    is_draft_final_generated = models.BooleanField(default=False)
    draft_final_generated_at = models.DateTimeField(null=True, blank=True)
    draft_final_user = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='draft_final_set')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'akta_pt'

    def __str__(self):
        return self.nama_pt


class Pengurus(models.Model):
    akta = models.ForeignKey(AktaPT, on_delete=models.CASCADE, related_name='pengurus_set')
    sebutan = models.CharField(max_length=20, blank=True)
    nama = models.CharField(max_length=255, blank=True)
    jabatan = models.CharField(max_length=100, blank=True)
    saham_persen = models.IntegerField(default=0)
    gelar_depan = models.CharField(max_length=50, blank=True)
    gelar_belakang = models.CharField(max_length=100, blank=True)
    gelar_singkat = models.CharField(max_length=20, blank=True)
    tpl = models.CharField(max_length=100, blank=True)
    tgl_lahir = models.DateField(null=True, blank=True)
    nik = models.CharField(max_length=20, blank=True)
    npwp = models.CharField(max_length=30, blank=True)
    pekerjaan = models.CharField(max_length=100, blank=True)
    email = models.CharField(max_length=100, blank=True, default='')
    no_hp = models.CharField(max_length=50, blank=True, default='')
    tipe_kab = models.CharField(max_length=20, blank=True)
    kab = models.CharField(max_length=100, blank=True)
    jalan = models.TextField(blank=True)
    rt = models.CharField(max_length=10, blank=True)
    rw = models.CharField(max_length=10, blank=True)
    desa = models.CharField(max_length=100, blank=True)
    kecamatan = models.CharField(max_length=100, blank=True)
    provinsi = models.CharField(max_length=100, blank=True)
    kategori_yayasan = models.CharField(max_length=50, blank=True, default='')
    is_pendiri = models.BooleanField(default=False)
    ktp_urls = models.JSONField(default=list, blank=True)
    npwp_urls = models.JSONField(default=list, blank=True)

    class Meta:
        db_table = 'pengurus'

    def __str__(self):
        return self.nama

    def to_dict(self):
        return {
            'id': self.id,
            'akta_id': self.akta_id,
            'sebutan': self.sebutan or '',
            'nama': self.nama or '',
            'jabatan': self.jabatan or '',
            'saham_persen': int(self.saham_persen or 0),
            'gelar_depan': self.gelar_depan or '',
            'gelar_belakang': self.gelar_belakang or '',
            'gelar_singkat': self.gelar_singkat or '',
            'tpl': self.tpl or '',
            'tgl_lahir': self.tgl_lahir.strftime('%Y-%m-%d') if self.tgl_lahir else '',
            'nik': self.nik or '',
            'npwp': self.npwp or '',
            'email': self.email or '',
            'no_hp': self.no_hp or '',
            'pekerjaan': self.pekerjaan or '',
            'tipe_kab': self.tipe_kab or '',
            'kab': self.kab or '',
            'jalan': self.jalan or '',
            'rt': self.rt or '',
            'rw': self.rw or '',
            'desa': self.desa or '',
            'kecamatan': self.kecamatan or '',
            'provinsi': self.provinsi or '',
            'kategori_yayasan': self.kategori_yayasan or '',
            'is_pendiri': self.is_pendiri,
            'ktp_urls': self.ktp_urls if isinstance(self.ktp_urls, list) else [],
            'npwp_urls': self.npwp_urls if isinstance(self.npwp_urls, list) else [],
        }


class CatatanAkta(models.Model):
    akta = models.ForeignKey(AktaPT, on_delete=models.CASCADE, related_name='catatan_set')
    isi = models.TextField()
    selesai = models.BooleanField(default=False)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'catatan_akta'

class DraftUserStatus(models.Model):
    akta = models.ForeignKey(AktaPT, on_delete=models.CASCADE, related_name='user_statuses')
    user = models.ForeignKey(User, on_delete=models.CASCADE)
    is_lengkap = models.BooleanField(default=False)
    is_revisi = models.BooleanField(default=False)  # True=edit draft, False=draf baru
    date_recorded = models.DateField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        unique_together = ('akta', 'user')
