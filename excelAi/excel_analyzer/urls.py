"""
URL configuration for excel_analyzer app.
"""
from django.urls import path
from .views import ListSheetsView, RowsPreviewView

app_name = 'excel_analyzer'

urlpatterns = [
    path('sheets/', ListSheetsView.as_view(), name='list_sheets'),
    path('rows/<uuid:file_id>/', RowsPreviewView.as_view(), name='rows_preview'),
]

