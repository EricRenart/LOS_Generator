import pytest
import openpyxl as opxl
import pandas as pd
import LOSGenerator
import os

# Path for test files
TEST_PATH = os.path.dirname(os.path.abspath(__file__)) + "\\tests\\"

class LOSGeneratorTestHelpers:

    # Test Helper Functions.
    # Creates a blank workbook in the test directory with optional LOS headers. Omit .xlsx in filename.
    def create_blank_workbook(self, filename='test', initial_headers=False):
        wb = LOSGenerator.create_workbook(initial_headers=initial_headers)
        if os.path.splitext(TEST_PATH)[1] == ".xlsx": # ensure filename does not contain .xlsx
            raise ValueError("Specified filename should not contain .xlsx.")
        wb.save(os.path.join(TEST_PATH, f"{filename}.xlsx"))
        return wb

    # Deletes all files in test folder.
    def clear_test_data(self):
        for file in os.listdir(TEST_PATH):
            path = os.path.join(TEST_PATH, file)
            if os.path.isfile(path):
                os.remove(path)

class LOSGeneratorTests:
    # Test Functions

    def test_workbook_create(self):
        print('test_workbook_create()')
        fname = 'test_workbook_create'
        LOSGeneratorTestHelpers.create_blank_workbook(TEST_PATH)
        assert os.path.exists(TEST_PATH+f"\\{fname}.xlsx")

    def test_workbook_create_with_headers(self):
        print('test_workbook_create_with_headers()')
        fname = 'test_workbook_create_with_headers'
        LOSGeneratorTestHelpers.create_blank_workbook(TEST_PATH, initial_headers=True)

        # Does workbook exist?
        assert os.path.exists(TEST_PATH+f"\\{fname}.xlsx")

        # Do headers exist in workbook?
        wb = opxl.load_workbook(TEST_PATH+f"\\{fname}.xlsx")
        ws = wb.active
