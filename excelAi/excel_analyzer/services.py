"""
Business logic for Excel file analysis.
"""
from openpyxl import load_workbook
from typing import Dict, Any, Union, List
from django.core.files.uploadedfile import UploadedFile


def list_excel_sheets(file: UploadedFile) -> Dict[str, Union[List[str], int]]:
    """
    Read an Excel file and return list of sheet names.
    
    Args:
        file: Django uploaded file object
        
    Returns:
        Dictionary containing list of sheet names
    """
    workbook = load_workbook(file, read_only=True)
    
    return {
        'sheets': workbook.sheetnames,
        'total_sheets': len(workbook.sheetnames)
    }

