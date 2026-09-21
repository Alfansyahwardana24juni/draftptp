"""
Entry point WSGI untuk Vercel Serverless Function.

Vercel memanggil variabel `app` di file ini untuk setiap request.
Folder proyek (tempat manage.py) perlu dimasukkan ke sys.path karena
fungsi ini dijalankan dari dalam direktori api/.
"""
import os
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'velloscript.settings')

from velloscript.wsgi import application  # noqa: E402

# Nama `app` wajib — ini yang dicari runtime Python Vercel.
app = application
