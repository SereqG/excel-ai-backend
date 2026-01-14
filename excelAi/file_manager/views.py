"""
Views for file manager app.
"""
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework.parsers import MultiPartParser, FormParser
from rest_framework import status
from django.http import FileResponse
from .services import (
    check_authentication,
    validate_upload_request,
    handle_upload_file,
    get_file_by_id,
    format_file_status_response,
    format_upload_response,
    validate_file_for_download,
)
from .exceptions import (
    FileSizeExceededError,
    UploadLimitExceededError,
    FileNotFoundError,
    FileNotReadyError,
    AuthenticationError,
    RequestValidationError,
)
from .tasks import process_spreadsheet_sheet


class FileUploadView(APIView):
    """
    API endpoint to upload a spreadsheet file.
    
    Expects:
    - POST request with multipart/form-data
    - 'file' field: Excel file (.xlsx or .xls)
    - 'sheet_name' field: Name of the sheet to extract
    
    Returns:
    - file_id: UUID of the uploaded file
    - status: Current status of the file
    """
    parser_classes = [MultiPartParser, FormParser]
    
    def post(self, request):
        """Handle file upload."""
        try:
            clerk_user_id = check_authentication(request)
            
            uploaded_file, sheet_name = validate_upload_request(request)
            
            file_record = handle_upload_file(
                clerk_user_id=clerk_user_id,
                uploaded_file=uploaded_file,
                sheet_name=sheet_name
            )
            
            process_spreadsheet_sheet.delay(str(file_record.file_id))
            
            return Response(
                format_upload_response(file_record),
                status=status.HTTP_201_CREATED
            )
        
        except AuthenticationError as e:
            return Response(
                {'error': str(e)},
                status=status.HTTP_401_UNAUTHORIZED
            )
        
        except RequestValidationError as e:
            return Response(
                {'error': str(e)},
                status=status.HTTP_400_BAD_REQUEST
            )
        
        except FileSizeExceededError as e:
            return Response(
                {'error': str(e)},
                status=status.HTTP_400_BAD_REQUEST
            )
        
        except UploadLimitExceededError as e:
            return Response(
                {'error': str(e)},
                status=status.HTTP_429_TOO_MANY_REQUESTS
            )
        
        except Exception as e:
            return Response(
                {'error': f'Error uploading file: {str(e)}'},
                status=status.HTTP_500_INTERNAL_SERVER_ERROR
            )


class FileStatusView(APIView):
    """
    API endpoint to check the status of a processed file.
    
    Expects:
    - GET request with file_id in URL
    
    Returns:
    - file_id: UUID of the file
    - status: Current status (UPLOADED, PROCESSING, READY, FAILED, EXPIRED)
    - created_at: When the file was uploaded
    - expires_at: When the file will expire
    """
    
    def get(self, request, file_id):
        """Get file status."""
        try:
            # Check authentication
            clerk_user_id = check_authentication(request)
            
            # Get file record
            file_record = get_file_by_id(file_id, clerk_user_id)
            
            # Return formatted response
            return Response(
                format_file_status_response(file_record),
                status=status.HTTP_200_OK
            )
        
        except AuthenticationError as e:
            return Response(
                {'error': str(e)},
                status=status.HTTP_401_UNAUTHORIZED
            )
        
        except FileNotFoundError as e:
            return Response(
                {'error': str(e)},
                status=status.HTTP_404_NOT_FOUND
            )
        
        except Exception as e:
            return Response(
                {'error': f'Error retrieving file status: {str(e)}'},
                status=status.HTTP_500_INTERNAL_SERVER_ERROR
            )


class FileDownloadView(APIView):
    """
    API endpoint to download a processed file.
    
    Expects:
    - GET request with file_id in URL
    
    Returns:
    - Processed Excel file as download
    """
    
    def get(self, request, file_id):
        """Download processed file."""
        try:
            # Check authentication
            clerk_user_id = check_authentication(request)
            
            # Get file record
            file_record = get_file_by_id(file_id, clerk_user_id)
            
            # Validate file is ready for download
            validate_file_for_download(file_record)
            
            # Return file as download
            response = FileResponse(
                file_record.processed_file.open('rb'),
                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            response['Content-Disposition'] = f'attachment; filename="{file_record.processed_file.name}"'
            
            return response
        
        except AuthenticationError as e:
            return Response(
                {'error': str(e)},
                status=status.HTTP_401_UNAUTHORIZED
            )
        
        except FileNotFoundError as e:
            return Response(
                {'error': str(e)},
                status=status.HTTP_404_NOT_FOUND
            )
        
        except FileNotReadyError as e:
            return Response(
                {'error': str(e)},
                status=status.HTTP_400_BAD_REQUEST
            )
        
        except Exception as e:
            return Response(
                {'error': f'Error downloading file: {str(e)}'},
                status=status.HTTP_500_INTERNAL_SERVER_ERROR
            )
