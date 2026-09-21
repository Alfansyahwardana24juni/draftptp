from django.urls import path

from . import views

urlpatterns = [
    path('', views.form, name='form'),
    path('riwayat/', views.riwayat, name='riwayat'),
    path('cari-kbli/', views.cari_kbli, name='cari_kbli'),
    path('generate/', views.generate, name='generate'),
]
