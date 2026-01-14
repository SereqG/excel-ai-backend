"""
Business logic for file_manager app.
"""
import os
from typing import Optional
from django.core.exceptions import ValidationError, PermissionDenied
from django.conf import settings
from django.utils import timezone
from datetime import timedelta
from .models import ProcessedFile
from .exceptions import (
    FileSizeExceededError,
    UploadLimitExceededError,
    FileNotFoundError,
    FileNotReadyError,
    AuthenticationError,
    RequestValidationError,
)


def validate_file_size(file) -> None:
    """
    Validate that file size is within the allowed limit.
    
    Args:
        file: Django uploaded file object
        
    Raises:
        FileSizeExceededError: If file size exceeds MAX_FILE_SIZE
    """
    if file.size > settings.MAX_FILE_SIZE:
        raise FileSizeExceededError(
            f'File size ({file.size} bytes) exceeds maximum allowed size '
            f'({settings.MAX_FILE_SIZE} bytes)'
        )


def check_upload_limit(clerk_user_id: str) -> bool:
    """
    Check if user has exceeded daily upload limit.
    
    Args:
        clerk_user_id: Clerk user ID
        
    Returns:
        True if under limit, False otherwise
        
    Raises:
        UploadLimitExceededError: If limit exceeded
    """
    today_start = timezone.now().replace(hour=0, minute=0, second=0, microsecond=0)
    
    upload_count = ProcessedFile.objects.filter(
        clerk_user_id=clerk_user_id,
        created_at__gte=today_start
    ).count()
    
    if upload_count >= settings.MAX_UPLOADS_PER_DAY:
        raise UploadLimitExceededError(
            f'Daily upload limit of {settings.MAX_UPLOADS_PER_DAY} files exceeded. '
            f'You have uploaded {upload_count} files today.'
        )
    
    return True


def create_file_record(clerk_user_id: str, file, sheet_name: str) -> ProcessedFile:
    """
    Create a new file record after validation.
    
    Args:
        clerk_user_id: Clerk user ID
        file: Django uploaded file object
        sheet_name: Name of the selected sheet
        
    Returns:
        ProcessedFile instance
        
    Raises:
        FileSizeExceededError: If file size exceeds limit
        UploadLimitExceededError: If upload limit exceeded
        ValidationError: For other validation errors
    """
    validate_file_size(file)
    
    check_upload_limit(clerk_user_id)
    
    expires_at = timezone.now() + timedelta(hours=settings.FILE_TTL_HOURS)
    
    processed_file = ProcessedFile.objects.create(
        clerk_user_id=clerk_user_id,
        original_file=file,
        selected_sheet=sheet_name,
        status='UPLOADED',
        expires_at=expires_at
    )
    
    return processed_file


def get_file_by_id(file_id: str, clerk_user_id: str) -> ProcessedFile:
    """
    Retrieve a file by ID and enforce ownership.
    
    Args:
        file_id: UUID of the file
        clerk_user_id: Clerk user ID
        
    Returns:
        ProcessedFile instance
        
    Raises:
        FileNotFoundError: If file not found
        PermissionDenied: If user doesn't own the file
    """
    try:
        file_obj = ProcessedFile.objects.get(file_id=file_id)
    except ProcessedFile.DoesNotExist:
        raise FileNotFoundError(f'File with ID {file_id} not found')
    
    # Enforce ownership
    if file_obj.clerk_user_id != clerk_user_id:
        raise PermissionDenied('You do not have permission to access this file')
    
    # Update last accessed time
    file_obj.last_accessed_at = timezone.now()
    file_obj.save(update_fields=['last_accessed_at'])
    
    return file_obj


def mark_file_expired(file_id: str) -> None:
    """
    Mark a file as expired.
    
    Args:
        file_id: UUID of the file
    """
    try:
        file_obj = ProcessedFile.objects.get(file_id=file_id)
        file_obj.status = 'EXPIRED'
        file_obj.save(update_fields=['status'])
    except ProcessedFile.DoesNotExist:
        pass  # File already deleted, ignore


def delete_file_from_disk(file_path: str) -> None:
    """
    Safely delete a file from disk.
    
    Args:
        file_path: Full path to the file
    """
    if file_path and os.path.exists(file_path):
        try:
            os.remove(file_path)
        except OSError:
            # File might be locked or already deleted, ignore
            pass


def check_authentication(request) -> str:
    """
    Check if request has valid Clerk user ID.
    
    Args:
        request: Django request object
        
    Returns:
        Clerk user ID string
        
    Raises:
        AuthenticationError: If authentication is missing
    """
    clerk_user_id = getattr(request, 'clerk_user_id', None)
    if not clerk_user_id:
        raise AuthenticationError('Authentication required')
    return clerk_user_id


def validate_file_type(file) -> None:
    """
    Validate that file is an Excel file.
    
    Args:
        file: Django uploaded file object
        
    Raises:
        RequestValidationError: If file type is invalid
    """
    if not file.name.endswith(('.xlsx', '.xls')):
        raise RequestValidationError('Invalid file type. Please upload an Excel file (.xlsx or .xls).')


def validate_upload_request(request) -> tuple:
    """
    Validate upload request and extract file and sheet name.
    
    Args:
        request: Django request object
        
    Returns:
        Tuple of (uploaded_file, sheet_name)
        
    Raises:
        RequestValidationError: If request is invalid
    """
    if 'file' not in request.FILES:
        raise RequestValidationError('No file provided. Please upload an Excel file.')
    
    if 'sheet_name' not in request.data:
        raise RequestValidationError('sheet_name is required')
    
    uploaded_file = request.FILES['file']
    sheet_name = request.data['sheet_name']
    
    validate_file_type(uploaded_file)
    
    return uploaded_file, sheet_name


def format_file_status_response(file_record: ProcessedFile) -> dict:
    """
    Format file status response data.
    
    Args:
        file_record: ProcessedFile instance
        
    Returns:
        Dictionary with file status information
    """
    return {
        'file_id': str(file_record.file_id),
        'status': file_record.status,
        'created_at': file_record.created_at.isoformat(),
        'expires_at': file_record.expires_at.isoformat(),
        'is_expired': file_record.is_expired,
    }


def format_upload_response(file_record: ProcessedFile) -> dict:
    """
    Format file upload response data.
    
    Args:
        file_record: ProcessedFile instance
        
    Returns:
        Dictionary with upload response information
    """
    return {
        'file_id': str(file_record.file_id),
        'status': file_record.status,
        'message': 'File uploaded successfully. Processing started.'
    }


def validate_file_for_download(file_record: ProcessedFile) -> None:
    """
    Validate that file is ready for download.
    
    Args:
        file_record: ProcessedFile instance
        
    Raises:
        FileNotReadyError: If file is not ready
        FileNotFoundError: If processed file doesn't exist
    """
    if file_record.status != 'READY':
        raise FileNotReadyError(
            f'File is not ready for download. Current status: {file_record.status}'
        )
    
    if not file_record.processed_file:
        raise FileNotFoundError('Processed file not found')


def handle_upload_file(clerk_user_id: str, uploaded_file, sheet_name: str) -> ProcessedFile:
    """
    Handle file upload process: validate and create file record.
    
    Args:
        clerk_user_id: Clerk user ID
        uploaded_file: Django uploaded file object
        sheet_name: Name of the selected sheet
        
    Returns:
        ProcessedFile instance
        
    Raises:
        FileSizeExceededError: If file size exceeds limit
        UploadLimitExceededError: If upload limit exceeded
        RequestValidationError: For other validation errors
    """
    validate_file_type(uploaded_file)
    
    file_record = create_file_record(
        clerk_user_id=clerk_user_id,
        file=uploaded_file,
        sheet_name=sheet_name
    )
    
    return file_record
