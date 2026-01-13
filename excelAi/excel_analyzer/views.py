"""
Views for Excel analyzer app.
"""
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework.parsers import MultiPartParser, FormParser
from rest_framework import status
from .services import list_excel_sheets


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
