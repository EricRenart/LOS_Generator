from io import StringIO
import pandas as pd
import openpyxl as opxl
from openpyxl.styles import Alignment, Border, PatternFill, Side
import tkinter as tk
from tkinter import filedialog

# Colors
lt_grey = "8e8e8e"
drk_grey = "474747"
black = "ffffff"
white = "000000"

# Create a row of LOS data in the specified worksheet
def _create_movement_row(ws, start_row, period, direction):

    # Create section header
    ws.cell(row=start_row, column=2, value=period).alignment = Alignment(horizontal='center')
    ws.cell(row=start_row, column=3, value=direction).alignment = Alignment(horizontal='center')

    # Copy headers from the top for data section
    for column in range(1,29):
        value = ws.cell(row=2, column=column).value
        ws.cell(row=start_row + 1, column=column, value=value).alignment = Alignment(horizontal='center')

    # Copy subheaders
    for column in range(1,29):
        value = ws.cell(row=3, column=column).value
        ws.cell(row=start_row + 2, column=column, value=value).alignment = Alignment(horizontal='center')

# Format cells with borders and typeface
def format_cell(ws, row, col, sides='lrtb', style='thick', color="000000"):
    top, bottom, left, right = None
    if 'l' in sides:
        left = Side(border_style=style, color=color)
    if 'r' in sides:
        right = Side(border_style=style, color=color)
    if 't' in sides:
        top = Side(border_style=style, color=color)
    if 'b' in sides:
        bottom = Side(border_style=style, color=color)

    # Create and Apply border
    border = Border(top=top, left=left, right=right, bottom=bottom)
    fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
    ws.cell(row=row, column=col).border = border
    ws.cell(row=row, column=col).fill = fill


# Fill LOS values in cells
def _fill_data(ws, row, col, data, sides='lrtb'):
    ws.cell(row=row, column=col).value = data.loc[2] # Volume
    ws.cell(row=row, column=col+1).value = data.loc[3] # Synchro delay
    ws.cell(row=row, column=col+2).value = data.loc[4] # SimTraffic delay
    ws.cell(row=row, column=col+3).value = data.loc[5] # Percent volume in simtraffic
    ws.cell(row=row, column=col+4).value = data.loc[6] # v/c
    ws.cell(row=row, column=col+5).value = data.loc[7] # LOS
    ws.cell(row=row, column=col+6).value = data.loc[8] # Synchro Q50
    ws.cell(row=row, column=col+7).value = data.loc[9] # Synchro Q95
    ws.cell(row=row, column=col+8).value = data.loc[10] # Simtraffic Q50
    ws.cell(row=row, column=col+9).value = data.loc[11] # Simtraffic Q95
    ws.cell(row=row, column=col+10).value = data.loc[12] # Cycle Length
    ws.cell(row=row, column=col+11).value = data.loc[13] # Split
    ws.cell(row=row, column=col+12).value = data.loc[14] # Offset
    ws.cell(row=row, column=col+13).value = data.loc[15] # Notes

    # Apply border
    for i in range(0,13):

        # Upper-left cell
        if row == 0 and col == 0:
            format_cell(ws, row, col=col+i, sides='lt', style='thick')

        # Upper-right cell
        elif row == 0 and col == 15:
            format_cell(ws, row, col=col+i, sides='tr', style='thick')

        # Top line
        elif row == 0:
            format_cell(ws, row, col=col+i, sides='t', style='thick')

        # Bottom-left cell
        elif row == len(data) and col == 0:
            format_cell(ws, row, col=col+i, sides='bl', style='thick')

        # Bottom-right cell
        elif row == len(data) and col == 15:
            format_cell(ws, row, col=col+i, sides='br', style='thick')

        # Bottom line
        elif row == len(data):
            format_cell(ws, row, col=col+i, sides='b', style='thick')

        # Left line
        elif col == 0:
            format_cell(ws, row, col=col+i, sides='l', style='thick')

        # Right line
        elif col == 15:
            format_cell(ws, row, col=col+i, sides='r', style='thick')

        # Inner cells
        else:
            format_cell(ws, row, col=col+i, sides='lrtb', style='thin')

    # Advance row
    row += 1

# Creates Excel workbook and sets up initial headers
def create_workbook(initial_headers=False):
    wb = opxl.Workbook()
    if initial_headers:
        write_header_rows(wb.active, at_row=0)
    return wb

def write_header_rows(ws, at_row):
    # Lists of Headers (Columns A-S)
    # Row 1
    headers_1 = ["Node","","Street Name","","EXISTING","","","","","","","","",
                 "Node","","Street Name","","","","","PROPOSED","","","","","","","",""]

    # Row 2
    headers_2 = ["Time","Direction","Mvmt","Link Dist","Volume","Delay","Delay","LOS","LOS","Vol%","v/c","Q50","Q95",
                 "Q95","Cycle Length", "Split","Offset","Notes", "", "Time","Direction","Mvmt","Link Dist","Volume",
                 "Delay","Delay","LOS","LOS","Vol%","v/c","Q50","Q95","Q95", "Cycle Length","Split","Offset","Notes"]

    # Row 3
    headers_3 = ["","","","","","Synchro","Synchro","SimT","Synchro","SimT","SimT","Synchro","Synchro","Synchro","SimT",
                 "","","","","Synchro","Synchro","SimT","Synchro","SimT","SimT","Synchro","Synchro","Synchro","SimT"]

    # Write line 1
    for i, header in enumerate(headers_1):
        ws.write(at_row, i, header)

    # Write line 2
    for i, header in enumerate(headers_2):
        ws.write(at_row+1, i, header)

    # Write line 3
    for i, header in enumerate(headers_3):
        ws.write(at_row+2, i, header)

# Populate Excel sheet with data from text file
def populate_xlsx(txt_path, xlsx_path):

    # Create workbook and select active sheet
    wb = opxl.Workbook()
    ws = wb.active

    # Create workbook with initial headers
    create_workbook(ws, xlsx_path)

    # Read in Synchro file
    data = pd.read_csv(txt_path, skipinitialspace=True, header=None)

    # Create a Pandas ExcelWriter using xlsxwriter
    xlsw = pd.ExcelWriter(xlsx_path, engine='xlsxwriter')
    #sheet = xlsw.cur_sheet()

    # Write data to xlsxwriter
    data.to_excel(xlsw, sheet_name='Signals Review Sheet', startrow=3, startcol=2, header=False, index=False)




# Main Script

# Open file chooser dialog for Synchro report .txt
# Create and hide Tk root window
root = tk.Tk()
root.withdraw()

# Ask user for input .txt file from Synchro
input_path = filedialog.askopenfilename(title="Select Synchro report .txt file", defaultextension=".txt",
                                        filetypes=(("Text files","*.txt"),("All files","*.*")))
print(f"Synchro report file: {input_path}")

# Hide tk again
root.deiconify()

# Ask user for output location with .xlsx filename
output_path = filedialog.asksaveasfilename(title="Save output as .xlsx", defaultextension=".xlsx",
                                               filetypes=(("Excel files","*.xlsx"),("All files","*.*")))
print(f"Output excel spreadsheet: {output_path}")

root.deiconify()

# Process the input data
populate_xlsx(input_path, output_path)