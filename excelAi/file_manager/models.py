"""
Models for file_manager app.
"""
import uuid
from django.db import models
from django.utils import timezone
from datetime import timedelta


class ProcessedFile(models.Model):
    """
    Model to track uploaded files and their processing lifecycle.
    """
    STATUS_CHOICES = [
        ('UPLOADED', 'Uploaded'),
        ('PROCESSING', 'Processing'),
        ('READY', 'Ready'),
        ('FAILED', 'Failed'),
        ('EXPIRED', 'Expired'),
    ]
    
    file_id = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    clerk_user_id = models.CharField(max_length=255, db_index=True)
    original_file = models.FileField(upload_to='temp/originals/')
    processed_file = models.FileField(upload_to='temp/processed/', null=True, blank=True)
    selected_sheet = models.CharField(max_length=255)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='UPLOADED')
    created_at = models.DateTimeField(auto_now_add=True)
    last_accessed_at = models.DateTimeField(auto_now=True)
    expires_at = models.DateTimeField()
    
    class Meta:
        ordering = ['-created_at']
        indexes = [
            models.Index(fields=['clerk_user_id', 'created_at']),
            models.Index(fields=['status', 'expires_at']),
        ]
    
    def __str__(self):
        return f"{self.file_id} - {self.clerk_user_id} - {self.status}"
    
    @property
    def is_expired(self):
        """Check if the file has expired."""
        return timezone.now() > self.expires_at
    
    def save(self, *args, **kwargs):
        """Override save to set expires_at if not set."""
        if not self.expires_at:
            self.expires_at = timezone.now() + timedelta(hours=24)
        super().save(*args, **kwargs)
