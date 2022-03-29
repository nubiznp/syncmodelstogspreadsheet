from unicodedata import name
from django.contrib import admin
from django.urls import path
from . import views
from django.views.generic import TemplateView

urlpatterns = [
    path('', views.Index.as_view(), name='index'),
    path('upload/', views.upload, name='upload'),
    path('export', views.export, name='export'),
    # path('main/', views.main, name='main'),


]