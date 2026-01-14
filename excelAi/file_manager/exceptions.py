"""
Custom exceptions for file_manager app.
"""


class FileSizeExceededError(Exception):
    """Raised when file size exceeds the maximum allowed size."""
    pass


class UploadLimitExceededError(Exception):
    """Raised when user has exceeded daily upload limit."""
    pass


class FileNotFoundError(Exception):
    """Raised when file is not found."""
    pass


class FileNotReadyError(Exception):
    """Raised when file is not ready for download."""
    pass


class AuthenticationError(Exception):
    """Raised when authentication is required but not provided."""
    pass


class RequestValidationError(Exception):
    """Raised when request validation fails."""
    pass

