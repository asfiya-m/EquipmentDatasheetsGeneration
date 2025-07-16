"""
populate_equipment_names.py
This script automates the population of equipment names in the Master Equipment datasheet file.
Created on Wed July 16
@author: AsfiyaKhanam
"""
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
import pandas as pd

def populate_equipment_names(master_file, streamtable_file):
    """
    Populate equipment names from streamtable into master datasheet.
    Writes equipment names into each matching sheet starting from D3 (row 3, columns D..).
    
    Args:
        master_file (BytesIO or path): Master datasheet file (.xlsx).
        streamtable_file (BytesIO or path): Detailed streamtable file (.xlsx).
        
    Returns:
        output (BytesIO): Updated Excel file.
        filename (str): Timestamped output filename.
        skipped (list): Equipment names that could not be matched.
    """

    # Load master workbook
    wb_master = load_workbook(master_file)

    # Read streamtable — equipment names from A4 down
    df_stream = pd.read_excel(streamtable_file, sheet_name="Equipment & Stream List", header=None)
    equipment_names = df_stream.iloc[3:, 0].dropna().astype(str).tolist()

    skipped = []

    # Get list of master sheet names (lowercase for matching)
    master_sheets = {sheet.lower(): sheet for sheet in wb_master.sheetnames}

    for equip in equipment_names:
        matched_sheet = None
        # Try to find sheet where sheet name is substring of equipment name
        for sheet_lc, sheet_orig in master_sheets.items():
            if sheet_lc in equip.lower():
                matched_sheet = sheet_orig
                break
        if matched_sheet:
            ws = wb_master[matched_sheet]
            # find the first empty column starting at D3 (row 3)
            col = 4  # column D
            while ws.cell(row=3, column=col).value is not None:
                col += 1
            ws.cell(row=3, column=col, value=equip)
        else:
            skipped.append(equip)

    # Save result
    output = BytesIO()
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    filename = f"Master_DataSheet_EquipmentPopulated_{timestamp}.xlsx"
    wb_master.save(output)
    output.seek(0)

    return output, filename, skipped
