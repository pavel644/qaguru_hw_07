import openpyxl, os
from zipfile import ZipFile
from PyPDF2 import PdfReader

resources_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources')
sample_files_dir = os.path.join(resources_dir, 'sample_files')
temp_files_dir = os.path.join(resources_dir, 'tmp_files')
test_zip_archive_name = 'sample.zip'

with ZipFile(os.path.join(temp_files_dir, test_zip_archive_name), "w") as myzip:
    for file in os.listdir(sample_files_dir):
        myzip.write(os.path.join(sample_files_dir, file), file)


def test_pdf_file_exist_in_archive():
    with ZipFile(os.path.join(temp_files_dir, test_zip_archive_name), "r") as myzip:
        assert 'sample_pdf.pdf' in myzip.namelist()


def test_txt_file_exist_in_archive():
    with ZipFile(os.path.join(temp_files_dir, test_zip_archive_name), "r") as myzip:
        assert 'sample_txt.txt' in myzip.namelist()


def test_xls_file_exist_in_archive():
    with ZipFile(os.path.join(temp_files_dir, test_zip_archive_name), "r") as myzip:
        assert 'sample_xls.xls' in myzip.namelist()


def test_xlsx_file_exist_in_archive():
    with ZipFile(os.path.join(temp_files_dir, test_zip_archive_name), "r") as myzip:
        assert 'sample_xlsx.xlsx' in myzip.namelist()
