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
        LOSGeneratorTestHelpers.clear_test_data()

    def test_workbook_create_with_headers(self):
        print('test_workbook_create_with_headers()')
        fname = 'test_workbook_create_with_headers'
        LOSGeneratorTestHelpers.create_blank_workbook(TEST_PATH, initial_headers=True)

        # Does workbook exist?
        assert os.path.exists(TEST_PATH.join(f"{fname}.xlsx"))

        # Do headers exist in workbook?
        wb = opxl.load_workbook(TEST_PATH.join(f"{fname}.xlsx"))
        ws = wb.active

        # Row 1
        valid_header_1 = ["Node", "", "Street Name", "", "EXISTING", "", "", "", "", "", "", "", "",
                     "Node", "", "Street Name", "", "", "", "", "PROPOSED", "", "", "", "", "", "", "", ""]

        # Row 2
        valid_header_2 = ["Time", "Direction", "Mvmt", "Link Dist", "Volume", "Delay", "Delay", "LOS", "LOS", "Vol%", "v/c",
                     "Q50", "Q95",
                     "Q95", "Cycle Length", "Split", "Offset", "Notes", "", "Time", "Direction", "Mvmt", "Link Dist",
                     "Volume",
                     "Delay", "Delay", "LOS", "LOS", "Vol%", "v/c", "Q50", "Q95", "Q95", "Cycle Length", "Split",
                     "Offset", "Notes"]

        # Row 3
        valid_header_3 = ["", "", "", "", "", "Synchro", "Synchro", "SimT", "Synchro", "SimT", "SimT", "Synchro", "Synchro",
                     "Synchro", "SimT",
                     "", "", "", "", "Synchro", "Synchro", "SimT", "Synchro", "SimT", "SimT", "Synchro", "Synchro",
                     "Synchro", "SimT"]

        # Check each row in the worksheet
        for row in ws.iter_rows(min_row=2, max_row=2, min_col=1, max_col=15):
            for cell in row:
                if cell.value not in valid_header_1:
                    # Fail if headers differ
                    pytest.fail()

        for row in ws.iter_rows(min_row=3, max_row=3, min_col=1, max_col=15):
            for cell in row:
                if cell.value not in valid_header_2:
                    # Fail if headers differ
                    pytest.fail()

        for row in ws.iter_rows(min_row=4, max_row=4, min_col=1, max_col=15):
            for cell in row:
                if cell.value not in valid_header_3:
                    # Fail if headers differ
                    pytest.fail()

        LOSGeneratorTestHelpers.clear_test_data()

    def test_synchro_file_import(self):
        xlsx_fname = "test_synchro_file_import"
        synchro_fname = "SampleSynchroImport"

        # Create a dataframe for Synchro data
        data = pd.read_csv(TEST_PATH.join(f"{synchro_fname}.txt"))

        # TBI

        LOSGeneratorTestHelpers.clear_test_data()

