"""
Pangkas nilai yang melebihi batas panjang kolom.

SQLite membiarkan nilai kepanjangan tersimpan, Postgres menolaknya. Jalankan
sekali terhadap database SQLite SEBELUM pindah ke Postgres, kalau tidak
`loaddata` akan gagal dengan:
    value too long for type character varying(N)

Pakai --cek untuk melihat saja tanpa mengubah apa pun:
    python perbaiki_panjang_data.py --cek
    python perbaiki_panjang_data.py
"""
import os
import sys

import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'velloscript.settings')
django.setup()

from core.models import AktaPT, Pengurus  # noqa: E402

CEK_SAJA = '--cek' in sys.argv


def main():
    total = 0
    for Model in (AktaPT, Pengurus):
        fields = [f for f in Model._meta.get_fields()
                  if getattr(f, 'max_length', None) and f.get_internal_type() == 'CharField']
        for obj in Model.objects.all():
            ubah = []
            for f in fields:
                val = getattr(obj, f.name)
                if isinstance(val, str) and len(val) > f.max_length:
                    ubah.append((f.name, len(val), f.max_length))
                    setattr(obj, f.name, val[:f.max_length])
            if ubah:
                total += len(ubah)
                rincian = ', '.join(f'{n} ({a}->{b})' for n, a, b in ubah)
                print(f"  {Model.__name__} pk={obj.pk}: {rincian}")
                if not CEK_SAJA:
                    obj.save(update_fields=[n for n, _, _ in ubah])

    if total == 0:
        print("Semua nilai sudah sesuai batas kolom. Tidak ada yang perlu diubah.")
    elif CEK_SAJA:
        print(f"\n{total} nilai melebihi batas. Jalankan tanpa --cek untuk memangkas.")
    else:
        print(f"\n{total} nilai dipangkas.")


if __name__ == '__main__':
    main()
