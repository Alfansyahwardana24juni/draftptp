# VELLOSCRIPT — Draft PT Perorangan

Aplikasi Django untuk membuat draft **Pernyataan Pendirian PT Perorangan** dan
mengunduhnya sebagai dokumen `.docx`.

## Fitur

- Form pengisian data perseroan, kegiatan usaha (KBLI), dan data pemilik
- Pencarian KBLI dari database lokal (`kbli.db`) + rekomendasi AI opsional
- Pengurai alamat otomatis: tempel satu alamat, kolom RT/RW/desa/kecamatan terisi
- Generate dokumen `.docx` dari template Word
- Riwayat draft dengan pencarian, catatan, edit, dan hapus
- Tema terang & gelap, tampilan menyesuaikan layar HP

## Jalankan di lokal

```bash
python -m venv .venv
.venv/bin/pip install -r requirements.txt

cp .env.example .env          # pastikan berisi DEBUG=True dan COOKIE_SECURE=False

.venv/bin/python manage.py migrate
.venv/bin/python manage.py createsuperuser
.venv/bin/python manage.py runserver
```

Buka http://127.0.0.1:8000 — akan diarahkan ke halaman masuk.

## Deploy

Lihat **[DEPLOY_VERCEL.md](DEPLOY_VERCEL.md)**. Ringkasnya: push ke GitHub, lalu
Import di Vercel dan isi environment variable — termasuk `DATABASE_URL` ke
Postgres, karena SQLite tidak bisa dipakai di serverless.

## Catatan

- Halaman admin (Django `/admin/` dan panel custom) sudah dihapus. Pengelolaan
  user lewat `manage.py createsuperuser`.
- `staticfiles/` ikut di-commit karena Vercel tidak menjalankan `collectstatic`.
  Setelah mengubah `static/`, jalankan `manage.py collectstatic --noinput --clear`.
