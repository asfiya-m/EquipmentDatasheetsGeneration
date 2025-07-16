"""

app.py

Streamlit frontend for automating the generation and population of a Master Equipment Datasheet.

Steps:
1. Upload a raw .xlsm file with multiple equipment sheets.
2. Generate a categorized master datasheet with grouped input sections.
3. Upload a SysCAD streamtable Excel file to populate SysCAD Inputs into the master datasheet.
4. Download final Excel file with all populated data.

Author: Asfiya Khanam
Created: June 2025

"""

import streamlit as st
from io import BytesIO
from datetime import datetime
import pandas as pd

# Backend code
from automation_test1 import generate_master_datasheet
from populate_equipment_names import populate_equipment_names

# from populate_syscad_inputs_rev1 import populate_syscad_inputs

st.title("📄 Master Equipment Datasheet Automation Tool")

st.markdown("""
This tool helps you:
1. Generate a clean, categorized master datasheet from the Excel datasheets workbook.
2. Populate Equipment names
""")

# ------------------------
# Step 1: Generate Master Sheet
# ------------------------
st.header("Step 1: Generate Master Datasheet")
st.markdown("""
**What happens in this step?**
- Extracts equipment-wise parameters from your datasheet file.
- Groups them under 5 categories:
    - SysCAD Inputs
    - Engineering Inputs
    - Lab/Pilot Inputs
    - Project Constants
    - Vendor Inputs
- Creates one formatted sheet per equipment.
""")

uploaded_raw = st.file_uploader("Upload your raw equipment .xlsm file", type=["xlsm"])
if uploaded_raw and st.button("Generate Master Sheet"):
    output_stream, output_filename = generate_master_datasheet(BytesIO(uploaded_raw.read()))
    output_stream.seek(0)  # Ensure it's at the beginning
    st.session_state["generated_master"] = output_stream

    st.success("✅ Master datasheet has been successfully generated!")

    st.download_button(
        label="📥 Download Master Sheet",
        data=output_stream,
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    
# ------------------------
# Step 2: Populate Equipment Names
# ------------------------
st.header("Step 2: Populate Equipment Names")
st.markdown("""
**What happens in this step?**
- Reads equipment names from your detailed streamtable.
- Looks for sheets in the master datasheet where the sheet name is a substring of the equipment name.
- Writes equipment names into the first available column starting at **D3**.
""")

# Check if Step 1 was completed
use_generated = False
if "generated_master" in st.session_state:
    use_generated = st.radio(
        "Choose master sheet to populate:",
        ["Use the one generated in Step 1", "Upload a different master sheet"]
    ) == "Use the one generated in Step 1"

if use_generated:
    master_bytes = st.session_state["generated_master"]
else:
    uploaded_master = st.file_uploader("Upload the master sheet", type=["xlsx"], key="master2")
    if uploaded_master:
        master_bytes = BytesIO(uploaded_master.read())
    else:
        master_bytes = None

uploaded_stream = st.file_uploader("Upload the detailed streamtable", type=["xlsx"], key="stream2")

if master_bytes and uploaded_stream and st.button("Populate Equipment Names"):
    stream_bytes = BytesIO(uploaded_stream.read())

    result, filename, skipped = populate_equipment_names(master_bytes, stream_bytes)

    if skipped:
        st.warning(f"⚠️ Some equipment were not matched to any sheet:\n{', '.join(skipped)}")

    st.success("✅ Equipment names populated successfully.")

    st.download_button(
        label="📥 Download Populated Master Sheet",
        data=result,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
