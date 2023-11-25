import openpyxl, os
from zipfile import ZipFile
from PyPDF2 import PdfReader
import pytest

resources_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources')
sample_files_dir = os.path.join(resources_dir, 'sample_files')
temp_files_dir = os.path.join(resources_dir, 'tmp_files')
test_zip_archive_name = 'sample.zip'

if not os.path.exists(temp_files_dir):
    os.mkdir(temp_files_dir)

with ZipFile(os.path.join(temp_files_dir, test_zip_archive_name), 'w') as myzip:
    for file in os.listdir(sample_files_dir):
        myzip.write(os.path.join(sample_files_dir, file), file)


def test_pdf_file_exist_in_archive():
    with ZipFile(os.path.join(temp_files_dir, test_zip_archive_name), 'r') as myzip:
        assert 'sample_pdf.pdf' in myzip.namelist()

def test_pdf_file_correct_size():
    with ZipFile(os.path.join(temp_files_dir, test_zip_archive_name), 'r') as myzip:
        assert myzip.getinfo('sample_pdf.pdf').file_size == 41540

def test_pdf_file_correct_data():
    with ZipFile(os.path.join(temp_files_dir, test_zip_archive_name), 'r') as myzip:
        with myzip.open('sample_pdf.pdf') as pdf_file:
            reader = PdfReader(pdf_file)
            assert len(reader.pages) == 2
            assert 'Text 123 @()#% ТестёжзфЩ' in reader.pages[0].extract_text()

def test_txt_file_exist_in_archive():
    with ZipFile(os.path.join(temp_files_dir, test_zip_archive_name), 'r') as myzip:
        assert 'sample_txt.txt' in myzip.namelist()

def test_txt_file_correct_size():
    with ZipFile(os.path.join(temp_files_dir, test_zip_archive_name), 'r') as myzip:
        assert myzip.getinfo('sample_txt.txt').file_size == 2265


def test_txt_file_correct_data():
    with ZipFile(os.path.join(temp_files_dir, test_zip_archive_name), 'r') as myzip:
        with myzip.open('sample_txt.txt') as txt_file:
            assert b'Quisque dictum faucibus risus' in txt_file.read()

def test_xls_file_exist_in_archive():
    with ZipFile(os.path.join(temp_files_dir, test_zip_archive_name), 'r') as myzip:
        assert 'sample_xls.xls' in myzip.namelist()


def test_xlsx_file_exist_in_archive():
    with ZipFile(os.path.join(temp_files_dir, test_zip_archive_name), 'r') as myzip:
        assert 'sample_xlsx.xlsx' in myzip.namelist()
