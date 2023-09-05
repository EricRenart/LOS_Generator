import pytest
import openpyxl as opxl
import pandas as pd
import LOSGenerator
import os

class LOSGeneratorTestData:
    PATH = os.path.dirname(os.path.abspath(__file__)) + "\\tests\\"

    # Data from an actual work project for testing input functionality
    NODE_NAMES = {1: "162nd Street & Northern Boulevard",
                     2: "162nd Street & Crocheron Avenue",
                     3: "Crocheron Avenue & 161st Street"}
    LANE_GROUPS = {1: ['EBL','EBT','EBR','WBL','WBT','WBR','NBL','NBT','NBR','SBL','SBT','SBR',
                       'SEL','SET','SER','NWL','NWT','NWR'],
                    2: ['EBL','EBT','EBR','WBL','WBT','WBR','NBL','NBT','NBR','SBL','SBT','SBR',
                    '�1','�5'],
                    3: ['EBL','EBT','EBR','WBL','WBT','WBR','NBL','NBT','NBR','SBL','SBT','SBR',
                    '�3','�7']}
    
    # HEaders for Excel sheet
    VALID_HEADERS = {1: ["Node", "", "Street Name", "", "EXISTING", "", "", "", "", "", "", "", "",
                    "Node", "", "Street Name", "", "", "", "", "PROPOSED", "", "", "", "", "", "", "", ""],
                     2: ["Time", "Direction", "Mvmt", "Link Dist", "Volume", "Delay", "Delay", "LOS", "LOS", "Vol%", "v/c",
                    "Q50", "Q95", "Q95", "Cycle Length", "Split", "Offset", "Notes", "", "Time", "Direction", "Mvmt", "Link Dist",
                    "Volume", "Delay", "Delay", "LOS", "LOS", "Vol%", "v/c", "Q50", "Q95", "Q95", "Cycle Length", "Split",
                    "Offset", "Notes"],
                     3: ["", "", "", "", "", "Synchro", "Synchro", "SimT", "Synchro", "SimT", "SimT", "Synchro", "Synchro",
                    "Synchro", "SimT",
                    "", "", "", "", "Synchro", "Synchro", "SimT", "Synchro", "SimT", "SimT", "Synchro", "Synchro",
                    "Synchro", "SimT"]}

class LOSGeneratorTestHelpers:

    # Test Helper Functions.
    # Creates a blank workbook in the test directory with optional LOS headers. Omit .xlsx in filename.
    def create_blank_workbook(self, filename='test', initial_headers=False):
        wb = LOSGenerator.create_workbook(initial_headers=initial_headers)
        if os.path.splitext(LOSGeneratorTestData.PATH)[1] == ".xlsx": # ensure filename does not contain .xlsx
            raise ValueError("Specified filename should not contain .xlsx.")
        wb.save(os.path.join(LOSGeneratorTestData.PATH, f"{filename}.xlsx"))
        return wb

    # Deletes all files in test folder.
    def clear_test_data(self):
        for file in os.listdir(LOSGeneratorTestData.PATH):
            path = os.path.join(LOSGeneratorTestData.PATH, file)
            if os.path.isfile(path):
                os.remove(path)

class LOSGeneratorTests:
    # Test Functions

    def test_workbook_create(self):
        print('test_workbook_create()')
        fname = 'test_workbook_create'
        LOSGeneratorTestHelpers.create_blank_workbook(LOSGeneratorTestData.PATH)
        assert os.path.exists(LOSGeneratorTestData.PATH.join(f"{fname}.xlsx")
        LOSGeneratorTestHelpers.clear_test_data()

    def test_workbook_create_with_headers(self):
        print('test_workbook_create_with_headers()')
        fname = 'test_workbook_create_with_headers'
        LOSGeneratorTestHelpers.create_blank_workbook(LOSGeneratorTestData.PATH, initial_headers=True)

        # Does workbook exist?
        assert os.path.exists(LOSGeneratorTestData.PATH.join(f"{fname}.xlsx"))

        # Do headers exist in workbook?
        wb = opxl.load_workbook(LOSGeneratorTestData.PATH.join(f"{fname}.xlsx"))
        ws = wb.active

        # Check each row in the worksheet
        for row in ws.iter_rows(min_row=2, max_row=2, min_col=1, max_col=15):
            for cell in row:
                if cell.value not in LOSGeneratorTestData.VALID_HEADERS[0]:
                    # Fail if headers differ
                    pytest.fail()

        for row in ws.iter_rows(min_row=3, max_row=3, min_col=1, max_col=15):
            for cell in row:
                if cell.value not in LOSGeneratorTestData.VALID_HEADERS[1]:
                    # Fail if headers differ
                    pytest.fail()

        for row in ws.iter_rows(min_row=4, max_row=4, min_col=1, max_col=15):
            for cell in row:
                if cell.value not in LOSGeneratorTestData.VALID_HEADERS[2]:
                    # Fail if headers differ
                    pytest.fail()

        LOSGeneratorTestHelpers.clear_test_data()

    def test_import_metadata_df(self):
        print('test_import_metadata_df()')
    
    def test_import_traffic_df(self):
        print('test_import_traffic_df()')
    
    def test_import_signal_df(self):
        print('test_import_signal_df()')
    
    def test_import_full(self):
        print('test_import_full()')
    
    def test_drop_empty_lane_groups(self):
        print('test_drop_empty_lane_groups()')

