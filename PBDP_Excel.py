# openpyxl Packages
from openpyxl import load_workbook as openpyxl_load_workbook
from openpyxl import Workbook as openpyxl_Workbook
from openpyxl.styles import Font as openpyxl_styles_Font

from typing import Union

# Open Excel File
def open_excel_file(filename: str) -> Union[openpyxl_Workbook, bool] :
    """Open an Excel file and return a workbook object."""
    try:
        wb = openpyxl_Workbook()
        wb = openpyxl_load_workbook(filename)
        return wb
    except:
        return False


# Create Excel File
def create_excel_file() -> Union[openpyxl_Workbook, bool] :
    """Create an Excel file"""
    try:
        wb = openpyxl_Workbook()
        ws = wb.active

        header = ['문제순서', '문제구분', '문제내용', '보기1', '보기2', '보기3', '보기4', '보기5', '정답']
        write_row_to_excel_file(wb, header, header=True)

        return wb
    except:
        return False


# Save Excel File
def save_excel_file(wb: openpyxl_Workbook, filename: str) -> bool :
    """Save an Excel file."""
    try:
        wb.save(filename)
        return True
    except:
        return False


# Write to Excel File
def write_row_to_excel_file(wb: openpyxl_Workbook, row: list, header=False) -> bool :
    """Write a row to an Excel file."""
    try:
        ws = wb.active
        ws.append(row)

        if header:
            header_row = ws[1]
            for cell in header_row:
                cell.font = openpyxl_styles_Font(bold=True, color= '000000FF')

        return True
    except:
        return False


# Read from Excel File
def read_rows_from_excel_file(wb: openpyxl_Workbook) -> Union[list, bool] :
    """Read rows from an Excel file."""
    try:
        ws = wb.active
        rows = []
        for row in ws.iter_rows():
            rows.append(row)
        return rows
    except:
        return False