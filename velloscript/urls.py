"""
URL configuration for velloscript project.
"""
from django.urls import path, include

urlpatterns = [
    path('', include('core.urls')),
]
