# EquipmentDatasheetsGeneration

This project automates the creation and population of a **Master Equipment Datasheet** from your SysCAD model outputs and standardized datasheet templates.  
It is designed as an internal Streamlit web app to make equipment documentation fast, consistent, and reliable.

## 📊 What does it do?
✅ Generates a clean, categorized **master equipment datasheet** from your `.xlsm` workbook.  
✅ Populates **equipment names** and **number of units** into the master sheet based on SysCAD outputs.  
✅ Populates **parameter values** into the master sheet by reading from `Stream Table V`, applying aggregation rules & unit conversions.  
✅ Uses an external YAML configuration file (`param_mapping.yaml`) for easy, flexible mapping of parameters → no hardcoding!

## 🧠 Key Features

- 📁 Categorized parameter sections:
  - SysCAD Inputs
  - Engineering Inputs
  - Lab/Pilot Inputs
  - Project Constants
  - Vendor Inputs

- 🧩 Rule-based parameter population:
  - Sum, average, or fixed values
  - Stream direction (input/output)
  - Stream overrides or fallbacks
  - Unit conversions (×1000, ÷100, etc.)

- ⚙️ YAML-driven configuration (`param_mapping.yaml`)
  - Per-equipment mappings
  - Column index, stream handling, conversion logic
  - Overrides for specific units

- 💡 Additional Features:
  - Equipment code mapping (e.g., `TK` → Tank)
  - Implied equipment generation (e.g., Agitators from Tanks)
  - Auto-counting number of units
  - ZIP export of individual equipment sheets (Step 4.5)
  - Optional verbose mode (backend only)

## App Workflow

### Step 1: Generate Master Datasheet
- Upload `.xlsm` file with multiple equipment sheets.
- App extracts parameters from each sheet.
- Categorizes and formats the content.
- Creates a multi-sheet Excel file:  
  `Master_DataSheet_<timestamp>.xlsx`

### Step 2: Populate Equipment Names
- Reads from `Equipment & Stream List` tab of SysCAD streamtable.
- Applies mapping logic to identify real and implied units.
- Writes equipment names in `D3` onward.
- Fills number of units at cell `B2`.

### Step 3: Populate Parameters
- Uses `Stream Table V` and YAML config to:
  - Resolve stream tags per unit
  - Apply column-specific value extraction
  - Apply rules for text, lookup, or numeric processing
- Handles override rules for specific units
- Logs skipped parameters if not found or invalid

### Step 4: Populate Engineering Inputs
- Uses column K of original datasheet workbook
- Pulls values for **Engineering Inputs** only
- Applies to all matching units
- Handles missing parameters gracefully

### Step 4.5: Export Individual Equipment Sheets
- Splits the final populated master workbook into **one file per sheet**
- Bundles them into a downloadable `.zip`
- Each file named after the sheet/equipment
- Available only after Step 4 is completed


### Developer Notes
- All processing is driven from a Streamlit app (app.py)
- Data is streamed between steps using st.session_state
- All file uploads are preserved between steps to minimize repetition
- Verbose mode is available internally (set verbose = True in code)

### Future Improvements
- Custom naming for split files in Step 4.5