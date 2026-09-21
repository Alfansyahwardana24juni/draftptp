# VELLOSCRIPT — Draft PT Perorangan

Aplikasi untuk membuat draft **Pernyataan Pendirian PT Perorangan** dan
mengunduhnya sebagai dokumen `.docx`.

## Tanpa database

Draft **disimpan di browser** masing-masing pengguna (`localStorage`).
Server tidak menyimpan apa pun — tugasnya hanya dua:

1. mencari kode KBLI dari `kbli.db` (dibaca saja), dan
2. merender `.docx` dari template Word memakai `docxtpl`, dan mengubahnya
   menjadi PDF bila diminta.

Rendering dokumen sengaja tetap di server: penanda seperti `{{ JLN_PT }}`
di dalam file Word terpecah ke beberapa bagian XML, dan `docxtpl`
menanganinya dengan benar. Implementasi di sisi browser rawan menghasilkan
dokumen rusak.

PDF disusun ulang dari dokumen .docx yang baru dirender (`core/pdf.py`),
lengkap dengan kop surat dari header Word. Ini penyusunan ulang, bukan
konversi piksel-per-piksel — konversi identik membutuhkan LibreOffice atau
Microsoft Word, yang tidak tersedia di serverless.

### Konsekuensi yang perlu diketahui

- Draft **tidak terbagi antar orang atau antar perangkat**.
- Membersihkan data browser, ganti laptop, atau memakai mode penyamaran
  berarti draft hilang.
- **Cadangkan secara berkala** lewat tombol **Ekspor** di halaman
  *Draft Tersimpan*, dan pulihkan lewat **Impor**.

## Deploy ke Vercel

Push ke GitHub, lalu di Vercel: **Add New → Project → Import**, pilih repo,
**Deploy**. Selesai.

Tidak ada Root Directory yang perlu diubah, tidak ada Build Command, dan
**tidak ada environment variable yang perlu diisi** — `vercel.json` sudah
mengatur semuanya.

## Jalankan di komputer sendiri

```bash
python -m venv .venv
.venv/bin/pip install -r requirements.txt

cp .env.example .env          # isinya cukup DEBUG=True

.venv/bin/python manage.py runserver
```

Buka http://127.0.0.1:8000 — tidak ada migrasi dan tidak ada pembuatan user,
karena memang tidak ada database.

## Halaman

| URL | Fungsi |
|---|---|
| `/` | Formulir; draft otomatis tersimpan ke browser sambil diketik |
| `/riwayat/` | Daftar draft tersimpan + Ekspor / Impor cadangan |
| `/cari-kbli/?q=` | Pencarian KBLI (JSON) |
| `/generate/` | POST data form → balas file `.docx` |
| `/generate-pdf/` | POST data form → balas file `.pdf` |

## Kalau mengubah file di static/

`staticfiles/` ikut di-commit karena Vercel tidak menjalankan
`collectstatic`. Setelah mengubah `static/`, jalankan sebelum push:

```bash
.venv/bin/python manage.py collectstatic --noinput --clear
```

Mengubah `templates/` tidak perlu langkah ini.
