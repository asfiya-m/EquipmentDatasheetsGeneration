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
            },
            "Solution Contents": {
                "col_idx": 2,          # Column B
                "agg": "first",
                "convert": None,
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

    # Iterate over all sheets & equipment names in master
    for sheet_name in wb_master.sheetnames:
        ws = wb_master[sheet_name]

        if sheet_name not in param_mapping:
            msg = f"[SKIP] Sheet '{sheet_name}': no param_mapping defined"
            skipped.append(msg)
            if verbose:
                print(msg)
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
                msg = f"[SKIP] {equip_name}: no streams found for base '{streams_key}'"
                if verbose: print(msg)
                skipped.append(msg)
                continue

            if verbose:
                print(f"\n✅ Equipment: {equip_name} → Sheet: {sheet_name} → Column: {equip_col}")

            collected = {k: [] for k in mapping_lc}

            for master_param, rule in mapping_lc.items():
                stream_type = rule.get("stream", "outlet")

                stream_tags = equip_stream_map[streams_key]["outputs"] if stream_type == "outlet" else equip_stream_map[streams_key]["inputs"]

                if not stream_tags:
                    msg = f"[SKIP] {equip_name}: no {stream_type} streams found"
                    if verbose: print(msg)
                    skipped.append(msg)
                    continue

                for stream_tag in stream_tags:
                    tag_lc = stream_tag.lower()
                    if tag_lc not in streamtable_lookup:
                        msg = f"[SKIP] {equip_name}: stream '{stream_tag}' not found"
                        if verbose: print(msg)
                        skipped.append(msg)
                        continue

                    stream_row = streamtable_lookup[tag_lc]
                    col_idx = rule["col_idx"]
                    try:
                        val = stream_row.iloc[col_idx - 1]
                        if rule["agg"] == "first":
                            if pd.notna(val) and not collected[master_param]:
                                collected[master_param].append(str(val))
                            continue

                        if pd.notna(val):
                            collected[master_param].append(float(val))

                        if verbose:
                            print(f" → Found {val} for {stream_type} stream {stream_tag}, param {master_param} (col {col_idx})")
                    except Exception as e:
                        msg = f"[SKIP] {equip_name}: failed reading {master_param} from stream '{stream_tag}': {e}"
                        if verbose: print(msg)
                        skipped.append(msg)

            # Write aggregated values
            for row_cells in ws.iter_rows(min_row=4, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                param_name = str(row_cells[1].value).strip() if row_cells[1].value else ""
                param_lc = param_name.lower()

                if param_lc in mapping_lc:
                    rule = mapping_lc[param_lc]
                    vals = collected.get(param_lc, [])

                    if not vals:
                        if verbose: print(f" → No values collected for {param_name}")
                        ws.cell(row=row_cells[0].row, column=equip_col).value = None
                        continue

                    if rule["agg"] == "first":
                        result = vals[0]
                    elif rule["agg"] == "sum":
                        result = np.nansum(vals)
                    elif rule["agg"] == "avg":
                        result = np.nanmean(vals)
                    else:
                        result = vals[0]

                    if rule["convert"] and isinstance(result, (int, float)):
                        result = rule["convert"](result)

                    if verbose:
                        print(f" → Writing {result} to cell ({row_cells[0].row},{equip_col}) for {param_name}")
                    ws.cell(row=row_cells[0].row, column=equip_col).value = result
                else:
                    ws.cell(row=row_cells[0].row, column=equip_col).value = None

    output = BytesIO()
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    filename = f"Master_DataSheet_ParametersPopulated_{timestamp}.xlsx"
    wb_master.save(output)
    output.seek(0)

    return output, filename, skipped
