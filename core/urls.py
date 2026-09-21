from django.urls import path
from django.views.generic import RedirectView
from . import views

urlpatterns = [
    path('', RedirectView.as_view(pattern_name='draft_pt_perorangan'), name='splash'),
    path('login/', views.login_view, name='login'),
    path('logout/', views.logout_view, name='logout'),

    path('draft/pt-perorangan/', views.draft_pt_perorangan, name='draft_pt_perorangan'),
    path('proses/', views.proses, name='proses'),
    path('proses_simpan/', views.proses_simpan, name='proses_simpan'),
    path('download/<int:akta_id>/', views.download, name='download'),

    path('riwayat/', views.riwayat, name='riwayat'),
    path('edit/<int:akta_id>/', views.edit, name='edit'),
    path('hapus/<int:akta_id>/', views.hapus, name='hapus'),

    path('cari_kbli/', views.cari_kbli, name='cari_kbli'),
    path('ai_kbli_recommendation/', views.get_kbli_recommendation, name='get_kbli_recommendation'),

    path('catatan/tambah/', views.catatan_tambah, name='catatan_tambah'),
    path('catatan/selesai/<int:catatan_id>/', views.catatan_selesai, name='catatan_selesai'),
]
