"""
populate_parameters.py

Step 3: Populates parameter values into the master equipment datasheet
based on stream data from the SysCAD streamtable.

Workflow:
1️⃣ Reads equipment names & stream tags from Equipment & Stream List sheet.
2️⃣ For each equipment, looks up associated streams.
3️⃣ For each stream, fetches parameter values from Stream Table V.
4️⃣ Aggregates & converts values as per defined rules.
5️⃣ Writes values back into the appropriate sheet & column in master sheet.

Author: Asfiya Khanam
Updated: July 2025
"""

from openpyxl import load_workbook
import pandas as pd
from io import BytesIO
from datetime import datetime
import numpy as np

def populate_parameters(master_file, streamtable_file, verbose=False):
    wb_master = load_workbook(master_file)

    # Read both sheets from streamtable file
    df_streamlist = pd.read_excel(
        streamtable_file, sheet_name="Equipment & Stream List", header=None, engine="openpyxl"
    )
    df_streamtable = pd.read_excel(
        streamtable_file, sheet_name="Stream Table V", header=6, engine="openpyxl"
    )

    skipped = []

    # Explicit mapping: equipment code → sheet name
    equipment_sheet_map = {
        "TK": "Tank",
        "FP_PK": "Filter Press",
        "IX_PK": "Ion Exchange",
        "RO_PK": "Reverse Osmosis System"
    }

    # Stream Table V lookup: stream tag → row
    streamtable_lookup = {}
    for idx, row in df_streamtable.iterrows():
        tag = str(row.iloc[0]).strip().lower()
        if tag:
            streamtable_lookup[tag] = row

    # Parameter mapping & aggregation rules
    param_mapping = {
        "Tank": {
            "Flow Rate to/from Vessel": {
                "col_idx": 7,   # Column G
                "agg": "sum",
                "convert": None
            },
            "Operating Temperature": {
                "col_idx": 10,  # Column J
                "agg": "avg",
                "convert": None
            },
            "Operating Density": {
                "col_idx": 15,  # Column O
                "agg": "avg",
                "convert": lambda x: x * 1000
            },
            "Design Density": {
                "col_idx": 15,  # Column O
                "agg": "avg",
                "convert": lambda x: x * 1000
            }
        }
    }

    # Loop over equipment rows in stream list
    equipment_rows = df_streamlist.iloc[3:]  # from row 4
    for _, row in equipment_rows.iterrows():
        equip_name = str(row[0]).strip()
        stream_tags = [str(tag).strip() for tag in row[1:5] if pd.notna(tag)]

        # Map equipment name → sheet
        matched_sheet = None
        for code, sheet_name in equipment_sheet_map.items():
            if equip_name.startswith(code):
                if sheet_name in wb_master.sheetnames:
                    matched_sheet = sheet_name
                else:
                    msg = f"[SKIP] {equip_name}: sheet '{sheet_name}' not found in master"
                    if verbose: print(msg)
                    skipped.append(msg)
                break

        if not matched_sheet:
            msg = f"[SKIP] {equip_name}: no matching sheet"
            if verbose: print(msg)
            skipped.append(msg)
            continue

        ws = wb_master[matched_sheet]
        equip_type = matched_sheet

        if equip_type not in param_mapping:
            msg = f"[SKIP] {equip_name}: no param mapping for sheet '{equip_type}'"
            if verbose: print(msg)
            skipped.append(msg)
            continue

        mapping = param_mapping[equip_type]
        mapping_lc = {k.strip().lower(): v for k, v in mapping.items()}
        collected = {k.strip().lower(): [] for k in mapping}

        # Find equipment column in master sheet (row 3)
        equip_col = None
        for cell in ws[3]:
            if cell.value and str(cell.value).strip().lower() == equip_name.strip().lower():
                equip_col = cell.column
                break

        if not equip_col:
            msg = f"[SKIP] {equip_name}: column not found in sheet '{matched_sheet}'"
            if verbose: print(msg)
            skipped.append(msg)
            continue

        if verbose:
            print(f"\n✅ Equipment: {equip_name} → Sheet: {matched_sheet} → Column: {equip_col}")

        # Collect values from associated streams
        for stream_tag in stream_tags:
            tag_lc = stream_tag.lower()
            if tag_lc not in streamtable_lookup:
                msg = f"[SKIP] {equip_name}: stream '{stream_tag}' not found"
                if verbose: print(msg)
                skipped.append(msg)
                continue

            stream_row = streamtable_lookup[tag_lc]

            # Extract parameter values per mapping
            for master_param, rule in mapping.items():
                param_lc = master_param.strip().lower()
                col_idx = rule["col_idx"]
                try:
                    val = stream_row.iloc[col_idx - 1]
                    if pd.notna(val):
                        collected[param_lc].append(float(val))
                    if verbose:
                        print(f" → Found value {val} for stream {stream_tag}, param {master_param} (col {col_idx})")
                except Exception as e:
                    msg = f"[SKIP] {equip_name}: failed reading {master_param} from stream '{stream_tag}': {e}"
                    if verbose: print(msg)
                    skipped.append(msg)

        # Write aggregated values back into master sheet
        for row_cells in ws.iter_rows(min_row=4, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            param_name = str(row_cells[1].value).strip() if row_cells[1].value else ""
            param_lc = param_name.lower()

            if param_lc in mapping_lc:
                rule = mapping_lc[param_lc]
                vals = collected.get(param_lc, [])

                if not vals:
                    if verbose: print(f" → No values collected for param: {param_name}")
                    ws.cell(row=row_cells[0].row, column=equip_col).value = None
                    continue

                # Apply aggregation
                if rule["agg"] == "sum":
                    result = np.nansum(vals)
                elif rule["agg"] == "avg":
                    result = np.nanmean(vals)
                else:
                    result = vals[0]

                # Apply conversion if needed
                if rule["convert"]:
                    result = rule["convert"](result)

                if verbose:
                    print(f" → Writing {round(result,2)} to cell ({row_cells[0].row},{equip_col}) for param {param_name}")
                ws.cell(row=row_cells[0].row, column=equip_col).value = round(result, 2)

            else:
                # No mapping for this parameter, clear value
                ws.cell(row=row_cells[0].row, column=equip_col).value = None

    # Save updated workbook
    output = BytesIO()
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    filename = f"Master_DataSheet_ParametersPopulated_{timestamp}.xlsx"
    wb_master.save(output)
    output.seek(0)

    return output, filename, skipped
