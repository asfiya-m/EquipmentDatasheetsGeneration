🚀 EquipmentDatasheetsGeneration
This project automates the creation and population of a Master Equipment Datasheet from your SysCAD model outputs and standardized datasheet templates.
It is designed as an internal Streamlit web app to make equipment documentation fast, consistent, and reliable.

📊 What does it do?
✅ Generates a clean, categorized master equipment datasheet from your .xlsm workbook.
✅ Populates equipment names and number of units into the master sheet based on SysCAD outputs.
✅ Populates parameter values into the master sheet by reading from Stream Table V, applying aggregation rules & unit conversions.
✅ Uses an external YAML configuration file (param_mapping.yaml) for easy, flexible mapping of parameters → no hardcoding!

📝 Features
✨ Generate a ready-to-use, properly formatted Master Equipment Datasheet with grouped categories:

SysCAD Inputs

Engineering Inputs

Lab/Pilot Inputs

Project Constants

Vendor Inputs

✨ Match and insert equipment names into the correct sheets & columns, counting the number of units.
✨ Populate parameter values like Flow, Temperature, Density, Pressure, etc., with rules:

sum / average

unit conversions (×1000, ÷1000, etc.)
✨ YAML-driven parameter mapping — easy to update & maintain.

✨ Optional verbose mode (backend-only) for detailed debugging.

🚀 Workflow Steps
📄 Step 1: Generate Master Datasheet
Upload a .xlsm workbook with equipment sheets.

App extracts parameters from each equipment sheet.

Groups parameters under standard categories.

Outputs one clean Excel file:

php-template
Copy
Edit
Master_DataSheet_<timestamp>.xlsx
🔷 Step 2: Populate Equipment Names
Reads equipment tags from Equipment & Stream List (in your SysCAD detailed streamtable).

Explicit mapping of codes:

pgsql
Copy
Edit
TK     → Tank
A      → Agitator (implied for each Tank)
FP_PK  → Filter Press
IX_PK  → Ion Exchange
RO_PK  → Reverse Osmosis System
Fills equipment names in master sheet starting at D3.

Adds number of units at B2.

Logs skipped equipment (if no matching sheet found).

📈 Step 3: Populate Parameters
Reads the master sheet (with equipment names) + the SysCAD streamtable.

Uses param_mapping.yaml for defining:

which parameters to populate

which streams to use (input/output)

which column in Stream Table V to read from

how to aggregate & convert

Supports:

Tank

Agitator

Filter Press

(Ion Exchange placeholder ready)

Writes values directly into the master sheet.

Logs skipped parameters & streams with clear messages when verbose=True.

🧾 Configuration
🔷 param_mapping.yaml
Defines all parameter mappings for each equipment type.

Example:

yaml
Copy
Edit
Tank:
  Operating Density:
    col_idx: 15
    agg: avg
    convert: multiply_1000
    stream: outlet

Filter Press:
  Feed material:
    stream_type: input
    stream_index: 0
    use_stream_name: true
  Solids density:
    col_idx: 16
    stream_type: input
    stream_index: 0
You can easily extend it with more parameters or equipment.

🐛 Debugging
✅ Run populate_parameters.py or populate_equipment_names.py with verbose=True for detailed logs:

python
Copy
Edit
result, filename, skipped = populate_parameters(master_bytes, stream_bytes, verbose=True)
Logs include:

which sheet & equipment matched

which streams were used

which parameter values were found & written

skipped items with clear reasons

In production → keep verbose=False.

📋 Potential Future Improvements
✨ Defensive checks for malformed equipment names.
✨ Automatic detection of YAML inconsistencies & validation.
✨ Build UI toggle for verbose mode.
✨ Add YAML mappings for Ion Exchange & Reverse Osmosis System.

📂 Files
app.py — Streamlit app frontend.

automation_test1.py — Step 1: generate master sheet.

populate_equipment_names.py — Step 2: populate equipment names & unit counts.

populate_parameters.py — Step 3: populate parameter values.

param_mapping.yaml — YAML file with all mappings.

requirements.txt — Python dependencies.

.gitignore — Ignore cache & unnecessary files.