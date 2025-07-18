"""
populate_parameters.py

Step 3: Populates parameter values into the master equipment datasheet
based on stream data from the SysCAD streamtable.

Handles both explicitly listed equipment and implied equipment (Agitator).
Maps parameters by equipment type using hardcoded column mappings and aggregation rules.

Workflow:
1️⃣ Reads equipment names & stream tags from Equipment & Stream List sheet.
2️⃣ For each equipment name in master file (including implied ones), finds corresponding input and output streams.
3️⃣ For each stream, fetches parameter values from Stream Table V, controlled by `param_mapping`.
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
                "col_idx": 7,          # Column G
                "agg": "sum",
                "convert": None,
                "stream": "outlet"
            },
            "Operating Temperature": {
                "col_idx": 10,         # Column J
                "agg": "avg",
                "convert": None,
                "stream": "outlet"
            },
            "Operating Density": {
                "col_idx": 15,         # Column O
                "agg": "avg",
                "convert": lambda x: x * 1000,
                "stream": "outlet"
            },
            "Design Density": {
                "col_idx": 15,         # Column O
                "agg": "avg",
                "convert": lambda x: x * 1000,
                "stream": "outlet"
            }
        },
        "Agitator": {
            "Flow Rate to/from Vessel": {
                "col_idx": 7,          # Column G
                "agg": "sum",
                "convert": None,
                "stream": "outlet"
            },
            "Operating Temperature": {
                "col_idx": 10,         # Column J
                "agg": "avg",
                "convert": None,
                "stream": "outlet"
            },
            "Operating Density": {
                "col_idx": 15,         # Column O
                "agg": "avg",
                "convert": lambda x: x * 1000,
                "stream": "outlet"
            },
            "Operating Pressure": {
                "col_idx": 11,         # Column K
                "agg": "avg",
                "convert": lambda x: x * 100,
                "stream": "outlet"
            }
        },
        "Filter Press": {
            "Cake Blow Required- Air Requirement": {
                "text": "N"
            },
            "Cake Wash Required- With What?- What Flow Rate?": {
                "text": "N"
            },
            "Feed material": {
                "stream_type": "input",
                "stream_index": 0,
                "use_stream_name": True
            },
            "Solids S.G.": {
                "col_idx": 17,                # Column Q
                "convert": lambda x: x / 1000,
                "stream_type": "input",
                "stream_index": 0
            },
            "Liquid SG": {
                "col_idx": 18,                # Column R
                "convert": lambda x: x / 1000,
                "stream_type": "input",
                "stream_index": 0
            },
            "Feed Solids Tonnage per Hour (Average)": {
                "col_idx": 4,                 # Column D
                "stream_type": "input",
                "stream_index": 0
            },
            "Feed Solids": {
                "col_idx": 12,                # Column L
                "stream_type": "input",
                "stream_index": 0
            },
            "Feed S.G. (t/m³)": {
                "col_idx": 15,               # Column O
                "stream_type": "input",
                "stream_index": 0
            },
            "Cake Solids Tonnage": {
                "col_idx": 6,                # Column F
                "stream_type": "output",
                "stream_index": 1
            },
            "Cake Moisture": {
                "col_idx": 18,               # Column R
                "stream_type": "output",
                "stream_index": 1
            },
            "Wet Cake Bulk Density": {
                "col_idx": 15,              # Column O
                "stream_type": "output",
                "stream_index": 1
            },
            "Filtrate Flow": {
                "col_idx": 7,               # Column G
                "stream_type": "output",
                "stream_index": 0
            }
        }
    }

    # Preprocess Equipment & Stream List into dict
    equip_stream_map = {}
    equipment_rows = df_streamlist.iloc[3:]  # from row 4
    for _, row in equipment_rows.iterrows():
        equip_name = str(row[0]).strip()
        outputs = [str(tag).strip() for tag in row[1:6] if pd.notna(tag)]  # B–F
        inputs  = [str(tag).strip() for tag in row[6:13] if pd.notna(tag)] # G–M
        equip_stream_map[equip_name] = {
            "outputs": outputs,
            "inputs": inputs
        }

    for sheet_name in wb_master.sheetnames:
        ws = wb_master[sheet_name]

        if sheet_name not in param_mapping:
            skipped.append(f"[SKIP] Sheet '{sheet_name}': no param_mapping defined")
            continue

        mapping = param_mapping[sheet_name]
        mapping_lc = {k.strip().lower(): v for k, v in mapping.items()}

        for col in range(4, ws.max_column + 1):
            cell = ws.cell(row=3, column=col)
            if not cell.value:
                continue
            equip_name = str(cell.value).strip()
            equip_col = cell.column

            streams_key = equip_name
            if sheet_name == "Agitator":
                if "-" not in equip_name:
                    skipped.append(f"[SKIP] {equip_name}: invalid format for implied equipment")
                    continue
                suffix = equip_name.split("-", 1)[1]
                streams_key = f"TK-{suffix}"

            if streams_key not in equip_stream_map:
                skipped.append(f"[SKIP] {equip_name}: no streams found for base '{streams_key}'")
                continue

            if verbose:
                print(f"\n✅ Equipment: {equip_name} → Sheet: {sheet_name} → Column: {equip_col}")

            for row_cells in ws.iter_rows(min_row=4, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                param_name = str(row_cells[1].value).strip() if row_cells[1].value else ""
                param_lc = param_name.lower()

                if param_lc not in mapping_lc:
                    ws.cell(row=row_cells[0].row, column=equip_col).value = None
                    continue

                rule = mapping_lc[param_lc]

                # text value
                if "text" in rule:
                    ws.cell(row=row_cells[0].row, column=equip_col).value = rule["text"]
                    if verbose:
                        print(f" → Writing text '{rule['text']}' to {param_name}")
                    continue

                # use stream name
                if rule.get("use_stream_name"):
                    stream_tags = equip_stream_map[streams_key][rule.get("stream_type", "output") + "s"]
                    idx = rule.get("stream_index", 0)
                    stream_name = stream_tags[idx] if idx < len(stream_tags) else ""
                    ws.cell(row=row_cells[0].row, column=equip_col).value = stream_name
                    if verbose:
                        print(f" → Writing stream name '{stream_name}' to {param_name}")
                    continue

                # numeric value
                stream_tags = equip_stream_map[streams_key][rule.get("stream_type", "output") + "s"]
                idx = rule.get("stream_index", 0)
                if idx >= len(stream_tags):
                    skipped.append(f"[SKIP] {equip_name}: no stream at index {idx} for {param_name}")
                    ws.cell(row=row_cells[0].row, column=equip_col).value = None
                    continue

                stream_tag = stream_tags[idx].lower()
                if stream_tag not in streamtable_lookup:
                    skipped.append(f"[SKIP] {equip_name}: stream '{stream_tag}' not found for {param_name}")
                    ws.cell(row=row_cells[0].row, column=equip_col).value = None
                    continue

                stream_row = streamtable_lookup[stream_tag]
                val = stream_row.iloc[rule["col_idx"] - 1]
                if pd.isna(val):
                    ws.cell(row=row_cells[0].row, column=equip_col).value = None
                    continue

                result = float(val)
                if rule.get("convert"):
                    result = rule["convert"](result)

                ws.cell(row=row_cells[0].row, column=equip_col).value = round(result, 2)
                if verbose:
                    print(f" → Writing {result} to {param_name}")

    output = BytesIO()
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    filename = f"Master_DataSheet_ParametersPopulated_{timestamp}.xlsx"
    wb_master.save(output)
    output.seek(0)

    return output, filename, skipped
