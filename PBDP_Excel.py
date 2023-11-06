# openpyxl Packages
from openpyxl import load_workbook as openpyxl_load_workbook
from openpyxl import Workbook as openpyxl_Workbook


# Open Excel File
def open_excel_file(filename: str) -> openpyxl_Workbook:
    """Open an Excel file and return a workbook object."""
    wb = openpyxl_Workbook()
    wb = openpyxl_load_workbook(filename)
    return wb


# Create Excel File
def create_excel_file():
    """Create an Excel file"""
    wb = openpyxl_Workbook()
    ws = wb.active
    header = ['문제순서', '문제구분', '문제내용', '보기1', '보기2', '보기3', '보기4', '보기5', '정답']
    ws.append(header)


    return wb


# Save Excel File
def save_excel_file(wb, filename: str) -> bool:
    """Save an Excel file."""
    try:
        wb.save(filename)
        return True
    except:
        return False