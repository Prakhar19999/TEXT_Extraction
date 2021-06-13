from django.contrib import admin
from django.urls import path
from .views import *

urlpatterns = [
    path('',pdf_to_text,name='home_page'),
    path(r'^export/xls/$', export_users_xls, name='export_users_xls'),
]