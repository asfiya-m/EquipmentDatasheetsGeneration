"""
populate_equipment_names.py
This script automates the population of equipment names in the Master Equipment datasheet file.
Created on Wed July 16
@author: AsfiyaKhanam
"""
from io import BytesIO
from datetime import datetime
from collections import defaultdict
import re
from openpyxl import load_workbook

# import pandas as pd

def populate_equipment_names(master_file, streamtable_file, verbose=True):
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
    ws_streamlist = load_workbook(streamtable_file, data_only=True)["Equipment & Stream List"]

    skipped = []

    # explicit mapping
    equipment_sheet_map = {
        "TK": "Tank",
        "BP_TK": "Bolted Panel Tank",
        "PF_TK": "PreFab Tank",
        "P_TK": "Poly Tank",
        "FP_PK": "Filter Press",
        "IX_PK": "Ion Exchange",
        "RO_PK": "Reverse Osmosis System",
        "S": "Clarifier",
        "E": "Heat Exchanger-1",
        "SL": "Silos",
        "F" : "Media Filter"

    }

    # implied mapping: for each TK → also create A in Agitator sheet
    implied_equipment = {
        "TK": [("A", "Agitator")]
    }

    #keep track of counts per sheet
    sheet_unit_counts = defaultdict(int)

     # read equipment names from streamtable
    equipment_names = []
    for row in ws_streamlist.iter_rows(min_row=4, min_col=1, max_col=1):
        cell_value = str(row[0].value).strip()
        if cell_value:
            equipment_names.append(cell_value)

    for equip_name in equipment_names:
        # 🚫 Skip if no numeric part
        match = re.search(r"-.*?(\d+)",equip_name)
        if not match:
            if verbose:
                print(f"⚠️ Skipping {equip_name}: no numeric part in name")
            skipped.append(f"{equip_name}: no numeric part in name")
            continue

        matched = False
        for code, sheet_name in equipment_sheet_map.items():
            equip_prefix = equip_name.split('-',1)[0]
            if equip_prefix == code:
            # if equip_name.startswith(code):
                if sheet_name in wb_master.sheetnames:
                    ws = wb_master[sheet_name]

                    # find first available column starting from D3
                    col_idx = 4  # D
                    while ws.cell(row=3, column=col_idx).value:
                        col_idx += 1

                    ws.cell(row=3, column=col_idx).value = equip_name
                    sheet_unit_counts[sheet_name] += 1

                    if verbose:
                        print(f"✅ Wrote {equip_name} → Sheet: {sheet_name} → Column: {col_idx}")
                    matched = True

                    # handle implied equipment if any
                    if code in implied_equipment:
                        number = equip_name.split("-")[1]
                        for implied_code, implied_sheet in implied_equipment[code]:
                            implied_name = f"{implied_code}-{number}"
                            if implied_sheet in wb_master.sheetnames:
                                ws_implied = wb_master[implied_sheet]
                                implied_col = 4
                                while ws_implied.cell(row=3, column=implied_col).value:
                                    implied_col += 1
                                ws_implied.cell(row=3, column=implied_col).value = implied_name
                                sheet_unit_counts[implied_sheet] += 1

                                if verbose:
                                    print(f"✨ Implied: Wrote {implied_name} → Sheet: {implied_sheet} → Column: {implied_col}")
                            else:
                                skipped.append(f"{implied_name}: sheet '{implied_sheet}' not found")
                                if verbose:
                                    print(f"⚠️ {implied_name}: sheet '{implied_sheet}' not found")

                    break
                else:
                    skipped.append(f"{equip_name}: sheet '{sheet_name}' not found")
                    if verbose:
                        print(f"⚠️ {equip_name}: sheet '{sheet_name}' not found")
                    matched = True
                    break

        if not matched:
            skipped.append(f"{equip_name}: no mapping for code")
            if verbose:
                print(f"⚠️ {equip_name}: no mapping for code")

    # write counts into B2 of each sheet
    for sheet_name, count in sheet_unit_counts.items():
        ws = wb_master[sheet_name]
        ws.cell(row=2, column=2).value = count  # B2
        if verbose:
            print(f"🔷 Sheet '{sheet_name}': unit count = {count} (written to B2)")

    # save output
    output = BytesIO()
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    filename = f"Master_DataSheet_EquipmentPopulated_{timestamp}.xlsx"
    wb_master.save(output)
    output.seek(0)

    return output, filename, skipped
