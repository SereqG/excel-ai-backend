"""
URL configuration for file_manager app.
"""
from django.urls import path
from .views import FileUploadView, FileStatusView, FileDownloadView

app_name = 'file_manager'

urlpatterns = [
    path('upload', FileUploadView.as_view(), name='file_upload'),
    path('status/<uuid:file_id>/', FileStatusView.as_view(), name='file_status'),
    path('download/<uuid:file_id>/', FileDownloadView.as_view(), name='file_download'),
]
