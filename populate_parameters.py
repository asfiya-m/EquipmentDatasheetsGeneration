from openpyxl import load_workbook
import pandas as pd
from io import BytesIO
from datetime import datetime
import numpy as np


def populate_parameters(master_file, streamtable_file, verbose=False):
    wb_master = load_workbook(master_file)
    df_streamlist = pd.read_excel(
        streamtable_file, sheet_name="Equipment & Stream List", header=None, engine="openpyxl"
    )
    df_streamtable = pd.read_excel(
        streamtable_file, sheet_name="Stream Table V", header=6, engine="openpyxl"
    )

    skipped = []

    # Build streamtable lookup
    streamtable_lookup = {}
    for idx, row in df_streamtable.iterrows():
        tag = str(row.iloc[0]).strip().lower()
        if tag:
            streamtable_lookup[tag] = row

    # updated density col_idx to 15
    param_mapping = {
        "Tank": {
            "Flow Rate to/from Vessel": {
                "col_idx": 7,  # G
                "agg": "sum",
                "convert": None
            },
            "Operating Temperature": {
                "col_idx": 10,  # J
                "agg": "avg",
                "convert": None
            },
            "Operating Density": {
                "col_idx": 15,  # O
                "agg": "avg",
                "convert": lambda x: x * 1000
            },
            "Design Density": {
                "col_idx": 15,  # O
                "agg": "avg",
                "convert": lambda x: x * 1000
            }
        }
    }

    equipment_rows = df_streamlist.iloc[3:]  # from row 4
    for _, row in equipment_rows.iterrows():
        equip_name = str(row[0]).strip()
        stream_tags = [str(tag).strip() for tag in row[1:5] if pd.notna(tag)]

        matched_sheet = None
        for sheet in wb_master.sheetnames:
            if sheet.lower() in equip_name.lower():
                matched_sheet = sheet
                break

        if not matched_sheet:
            msg = f"[SKIP] {equip_name}: sheet not found"
            if verbose: print(msg)
            skipped.append(msg)
            continue

        ws = wb_master[matched_sheet]
        equip_type = matched_sheet
        if equip_type not in param_mapping:
            msg = f"[SKIP] {equip_name}: no param mapping"
            if verbose: print(msg)
            skipped.append(msg)
            continue

        mapping = param_mapping[equip_type]
        mapping_lc = {k.strip().lower(): v for k, v in mapping.items()}
        collected = {k.strip().lower(): [] for k in mapping}

        equip_col = None
        for cell in ws[3]:
            if cell.value and str(cell.value).strip().lower() == equip_name.strip().lower():
                equip_col = cell.column
                break

        if not equip_col:
            msg = f"[SKIP] {equip_name}: column not found"
            if verbose: print(msg)
            skipped.append(msg)
            continue

        if verbose:
            print(f"\n✅ Equipment: {equip_name} → Sheet: {matched_sheet} → Column: {equip_col}")

        for stream_tag in stream_tags:
            tag_lc = stream_tag.lower()
            if tag_lc not in streamtable_lookup:
                msg = f"[SKIP] {equip_name}: stream '{stream_tag}' not found"
                if verbose: print(msg)
                skipped.append(msg)
                continue

            stream_row = streamtable_lookup[tag_lc]

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

        # Writing loop
        for row_cells in ws.iter_rows(min_row=4, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            category = str(row_cells[0].value).strip() if row_cells[0].value else ""
            param_name = str(row_cells[1].value).strip() if row_cells[1].value else ""

            if verbose:
                print(f"Row {row_cells[0].row}: Category='{category}', Param='{param_name}'")

            param_lc = param_name.strip().lower()
            if param_lc in mapping_lc:
                rule = mapping_lc[param_lc]
                vals = collected.get(param_lc, [])

                if not vals:
                    if verbose: print(f" → No values collected for param: {param_name}")
                    ws.cell(row=row_cells[0].row, column=equip_col).value = None
                    continue

                if rule["agg"] == "sum":
                    result = np.nansum(vals)
                elif rule["agg"] == "avg":
                    result = np.nanmean(vals)
                else:
                    result = vals[0]

                if rule["convert"]:
                    result = rule["convert"](result)

                if verbose:
                    print(f" → Writing {round(result,2)} to cell ({row_cells[0].row},{equip_col}) for param {param_name}")
                ws.cell(row=row_cells[0].row, column=equip_col).value = round(result, 2)

            else:
                ws.cell(row=row_cells[0].row, column=equip_col).value = None

    output = BytesIO()
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    filename = f"Master_DataSheet_ParametersPopulated_{timestamp}.xlsx"
    wb_master.save(output)
    output.seek(0)

    return output, filename, skipped
