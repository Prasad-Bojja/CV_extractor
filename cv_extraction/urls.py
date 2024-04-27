from django.urls import path
from cv_extraction.views import *

urlpatterns = [
    path('', upload_cv, name='upload_cv'),
    path('download-excel/<str:file_name>/', download_excel, name='download_excel'),
]
