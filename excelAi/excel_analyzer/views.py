"""
Views for Excel analyzer app.
"""
from django.core.exceptions import PermissionDenied
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework.parsers import MultiPartParser, FormParser
from rest_framework import status

from file_manager.exceptions import AuthenticationError, FileNotFoundError
from file_manager.services import check_authentication, get_file_by_id

from .services import list_excel_sheets
from .rows_preview import parse_truthy_query_flag, preview_rows_for_processed_file


class ListSheetsView(APIView):
    """
    API endpoint to analyze Excel file and return columns with types for each sheet.
    
    Expects a POST request with 'file' field containing the Excel file.
    Returns JSON response with sheets, their columns, and column types.
    """
    parser_classes = [MultiPartParser, FormParser]
    
    def post(self, request):
        """
        Handle POST request to analyze Excel file and return columns with types.
        """
        if 'file' not in request.FILES:
            return Response(
                {'error': 'No file provided. Please upload an Excel file.'},
                status=status.HTTP_400_BAD_REQUEST
            )
        
        uploaded_file = request.FILES['file']
        
        if not uploaded_file.name.endswith(('.xlsx', '.xls')):
            return Response(
                {'error': 'Invalid file type. Please upload an Excel file (.xlsx or .xls).'},
                status=status.HTTP_400_BAD_REQUEST
            )
        
        try:
            result = list_excel_sheets(uploaded_file)
            
            return Response(result, status=status.HTTP_200_OK)
        
        except Exception as e:
            return Response(
                {'error': f'Error processing Excel file: {str(e)}'},
                status=status.HTTP_500_INTERNAL_SERVER_ERROR
            )


class RowsPreviewView(APIView):
    """
    API endpoint to return first 10 rows (default) or random 10 rows of a user's uploaded sheet.

    GET /api/analyzer/rows/<file_id>/?random=true
    """

    def get(self, request, file_id):
        try:
            clerk_user_id = check_authentication(request)
            file_record = get_file_by_id(str(file_id), clerk_user_id)

            random_sample = parse_truthy_query_flag(request.query_params.get("random"))
            mode = "random" if random_sample else "first"

            preview = preview_rows_for_processed_file(
                original_path=file_record.original_file.path,
                selected_sheet=file_record.selected_sheet,
                random_sample=random_sample,
                limit=10,
            )

            return Response(
                {
                    "file_id": str(file_record.file_id),
                    "sheet_name": preview["sheet_name"],
                    "mode": mode,
                    "header": preview["header"],
                    "rows": preview["rows"],
                    "total_rows": preview["total_rows"],
                },
                status=status.HTTP_200_OK,
            )

        except AuthenticationError as e:
            return Response({"error": str(e)}, status=status.HTTP_401_UNAUTHORIZED)
        except FileNotFoundError as e:
            return Response({"error": str(e)}, status=status.HTTP_404_NOT_FOUND)
        except PermissionDenied as e:
            return Response({"error": str(e)}, status=status.HTTP_403_FORBIDDEN)
        except ValueError as e:
            return Response({"error": str(e)}, status=status.HTTP_400_BAD_REQUEST)
        except Exception as e:
            return Response(
                {"error": f"Error retrieving rows preview: {str(e)}"},
                status=status.HTTP_500_INTERNAL_SERVER_ERROR,
            )
