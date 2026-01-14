"""
Admin configuration for file_manager app.
"""
from django.contrib import admin
from .models import ProcessedFile


@admin.register(ProcessedFile)
class ProcessedFileAdmin(admin.ModelAdmin):
    """
    Admin interface for ProcessedFile model.
    """
    list_display = [
        'file_id',
        'clerk_user_id',
        'selected_sheet',
        'status',
        'created_at',
        'expires_at',
        'is_expired',
    ]
    
    list_filter = [
        'status',
        'created_at',
        'expires_at',
    ]
    
    search_fields = [
        'file_id',
        'clerk_user_id',
        'selected_sheet',
    ]
    
    readonly_fields = [
        'file_id',
        'created_at',
        'last_accessed_at',
        'expires_at',
    ]
    
    fieldsets = (
        ('File Information', {
            'fields': ('file_id', 'clerk_user_id', 'selected_sheet', 'status')
        }),
        ('Files', {
            'fields': ('original_file', 'processed_file')
        }),
        ('Timestamps', {
            'fields': ('created_at', 'last_accessed_at', 'expires_at')
        }),
    )
    
    ordering = ['-created_at']
