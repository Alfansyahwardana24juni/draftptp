# Deploy VELLOSCRIPT ke Vercel

## Tech stack tidak diganti

Vercel menjalankan Python sebagai Serverless Function, jadi Django ini dideploy
apa adanya — tidak ada fungsi yang berubah dan tidak ada yang ditulis ulang.
Isi proyek sudah dinaikkan ke root repo, jadi di Vercel cukup **Import** tanpa
mengubah Root Directory atau Build Command.

Satu hal infrastruktur yang wajib berubah: **database tidak boleh file SQLite
lagi, harus Postgres.** Filesystem serverless read-only dan ephemeral, jadi
`db.sqlite3` tidak bisa ditulis dan isinya hilang setiap deploy. `kbli.db`
tetap aman karena hanya dibaca (dibuka dengan `mode=ro`).

## Struktur repo

```
manage.py              vercel.json          requirements.txt
api/index.py           .vercelignore        .gitignore
core/                  velloscript/         templates/
static/                staticfiles/         template_word/
kbli.db                DEPLOY_VERCEL.md     perbaiki_panjang_data.py
```

| File | Fungsi |
|---|---|
| `api/index.py` | Entry point WSGI yang dipanggil Vercel (variabel `app`) |
| `vercel.json` | Routing semua path ke fungsi + daftar file yang dibundel |
| `requirements.txt` | Dependensi produksi (termasuk `psycopg` untuk Postgres) |
| `.vercelignore` | Menahan backup, `db.sqlite3`, dan `static/` agar tidak terunggah |
| `staticfiles/` | **Ikut di-commit** — lihat catatan di bawah |

## Langkah deploy

### 1. Siapkan Postgres

Pilih salah satu (semuanya punya paket gratis): **Neon**, **Supabase**, atau
**Vercel Postgres**. Ambil connection string-nya:

```
postgresql://user:password@host/dbname?sslmode=require
```

Kalau tersedia, pakai URL **pooled / pgbouncer** — lebih cocok untuk serverless.

### 2. Rapikan data yang melebihi batas kolom (WAJIB)

SQLite membiarkan nilai lebih panjang dari `max_length` tersimpan, Postgres
menolaknya. Kalau dilewati, `loaddata` gagal dengan
`value too long for type character varying(N)`.

```bash
.venv/bin/python perbaiki_panjang_data.py --cek   # lihat dulu
.venv/bin/python perbaiki_panjang_data.py         # pangkas
```

Penyebabnya sudah ditutup di `simpan_ke_db()` dan lewat atribut `maxlength`
di form, jadi data baru tidak akan kena lagi.

### 3. Migrasi + pindahkan data (dari komputer Anda)

Serverless tidak bisa menjalankan migrasi, jadi lakukan dari lokal:

```bash
export DATABASE_URL='postgresql://...'

.venv/bin/python manage.py migrate

.venv/bin/python manage.py dumpdata \
  --natural-foreign --natural-primary \
  -e contenttypes -e auth.Permission -e sessions \
  --indent 2 -o data_lama.json

DATABASE_URL="$DATABASE_URL" .venv/bin/python manage.py loaddata data_lama.json

.venv/bin/python manage.py createsuperuser
```

Hapus `data_lama.json` setelah selesai — isinya data pribadi (NIK, NPWP,
alamat). File itu sudah ada di `.gitignore`, tapi tetap sebaiknya dihapus.

### 4. Push ke GitHub

```bash
git remote add origin git@github.com:USERNAME/NAMA-REPO.git
git branch -M main
git push -u origin main
```

### 5. Import di Vercel

1. Vercel → **Add New → Project → Import Git Repository**
2. Pilih repo-nya. **Jangan ubah** Framework Preset, Root Directory, Build
   Command, atau Output Directory — `vercel.json` sudah mengatur semuanya.
3. Buka **Environment Variables** dan isi:

| Nama | Nilai |
|---|---|
| `SECRET_KEY` | string acak minimal 50 karakter |
| `DEBUG` | `False` |
| `COOKIE_SECURE` | `True` |
| `ALLOWED_HOSTS` | `.vercel.app` (tambah domain sendiri kalau ada) |
| `DATABASE_URL` | connection string Postgres dari langkah 1 |
| `DB_CONN_MAX_AGE` | `0` |
| `GEMINI_API_KEY` | opsional, untuk rekomendasi KBLI AI |

Kalau pakai domain sendiri, tambahkan juga
`CSRF_TRUSTED_ORIGINS=https://domain-anda.com`.

4. **Deploy.** Selanjutnya setiap `git push` ke `main` otomatis dideploy.

Kalau `DATABASE_URL` lupa diisi, aplikasi langsung berhenti dengan pesan jelas
("DATABASE_URL belum diset") daripada error yang membingungkan.

## Kalau mengubah file di static/

`staticfiles/` **sengaja ikut di-commit** karena Vercel tidak menjalankan
`collectstatic`. Setelah mengubah apa pun di `static/`, jalankan ini sebelum
push:

```bash
.venv/bin/python manage.py collectstatic --noinput --clear
git add staticfiles && git commit -m "Perbarui static"
```

Mengubah file di `templates/` tidak perlu langkah ini.

## Perubahan pada settings.py

- `DEBUG` default **False**; lokal diaktifkan lewat `.env`
- `ALLOWED_HOSTS` dari env, default `.vercel.app`; saat `DEBUG=True` jadi `['*']`
- `CSRF_TRUSTED_ORIGINS` sudah memuat `https://*.vercel.app`
- `SECURE_PROXY_SSL_HEADER` — wajib, TLS diakhiri di proxy Vercel
- `load_dotenv(override=False)` — environment variable platform menang atas `.env`
- `DB_CONN_MAX_AGE` default `0` — serverless tidak boleh menahan koneksi
- `STORAGES['staticfiles']` menggantikan `STATICFILES_STORAGE` yang dihapus
  Django 5.1 (sebelumnya diabaikan diam-diam, WhiteNoise tidak pernah aktif)
- Penjaga: berhenti dengan pesan jelas kalau produksi tanpa `DATABASE_URL`

## Sudah diuji lokal

Jalur produksi diuji dengan Postgres sementara (container dibuang setelahnya),
lewat `api/index.py`, `DEBUG=False`, dan header `X-Forwarded-Proto: https`:

- `migrate` ke Postgres: seluruh migrasi terpasang
- `dumpdata` + `loaddata`: 159 objek (45 akta, 45 pengurus, 15 user) masuk utuh
- Login, form, riwayat, edit: 200
- Generate .docx: 200, file .docx valid ±86 KB
- Pencarian KBLI dari `kbli.db` read-only: 200
- Static (CSS + gambar) lewat WhiteNoise: 200 dengan Content-Type benar
- Host di luar `ALLOWED_HOSTS`: ditolak 400

Yang **belum** diuji karena butuh akun Anda: proses import dan build di
infrastruktur Vercel itu sendiri (cold start, ukuran bundle, routing).

## Batasan Vercel

- **Durasi request**: 10 detik (Hobby) / 60 detik (Pro). Generate .docx ±0,2 detik.
- **Cold start**: request pertama setelah idle ±1–3 detik.
- **Ukuran bundle**: batas 250 MB uncompressed; bundle ini ±40 MB.
- **Tidak ada penyimpanan file permanen.** Aplikasi ini tidak menyimpan file
  (dokumen dikirim langsung dari memori), jadi tidak masalah.
