"""
Working with file
"""
import os.path
from zipfile import ZipFile
import zipfile
from PyPDF2 import PdfReader
import csv
from openpyxl import load_workbook

from config import folder_path, zip_name, pdf_text, names, pdf_size, marks


def file_to_zip(folder_with_file: str, zip_file_name: str):
    """
    Compress files in folder
    """
    zip_file = ZipFile(zip_file_name, 'w')

    for filename in os.listdir(folder_with_file):
        zip_file.write(os.path.join(folder_with_file, filename), compress_type=zipfile.ZIP_DEFLATED)

    return zip_file


def test_files():
    """
    Testing files
    """
    zip_file = file_to_zip(folder_path, zip_name)

    for file in zip_file.namelist():
        if file.endswith('.pdf'):
            pdf_reader = PdfReader(file)
            assert pdf_size, len(pdf_reader.pages)
            assert pdf_text in pdf_reader.pages[0].extractText()
        elif file.endswith('.csv'):
            with open(file, newline='') as csvfile:
                csv_reader = csv.DictReader(csvfile)
                for row in csv_reader:
                    assert row['name'] in names, f'Value must be one of {names}'
        elif file.endswith('.xlsx'):
            xlsx_reader = load_workbook(file)
            sheet = xlsx_reader.active
            for row, cell in enumerate(sheet['C']):
                if row != 0:
                    assert cell.value in marks, f'Value must be one of {marks}'
        else:
            print('File not supported...')

    zip_file.close()
