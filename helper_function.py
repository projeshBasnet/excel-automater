import re

def find_excel_sheet(each):
    """
    find the excel file for the student folder and return it
    """
    for file in each.rglob("*.xlsx"):        
        return file


def find_pdf_file(each):
    """
    find the pdf file for the student folder and return it
    """
    for file in each.rglob("*.pdf"):
        return file


def iterate_student_folder(folder_path):
    pass