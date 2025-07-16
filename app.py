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

# Backend code
from automation_test1 import generate_master_datasheet
# from populate_syscad_inputs_rev1 import populate_syscad_inputs

st.title("📄 Master Equipment Datasheet Automation Tool")

st.markdown("""
This tool helps you:
1. Generate a clean, categorized master datasheet from the Excel datasheets workbook.
2. Populate the master sheet with SysCAD streamtable data.

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
# Step 2: Populate with SysCAD Streamtable
# ------------------------
# st.header("Step 2: Populate with SysCAD Inputs")
# st.markdown("""
# **What happens in this step?**
# - Reads your generated master sheet (from Step 1 or manual upload).
# - Compares equipment sheets with those in the SysCAD streamtable.
# - Populates matching parameters in the **SysCAD Inputs** category.
# - Rounds values to two decimals.
# - Replaces the Master datasheet units with the ones from the streamtable if they differ.
# """)

# uploaded_master = st.file_uploader("Upload the master sheet (optional)", type=["xlsx"], key="master")
# uploaded_stream = st.file_uploader("Upload the SysCAD streamtable Excel", type=["xlsx"], key="stream")

# if st.button("Populate SysCAD Inputs"):
#     master_bytes = (
#         st.session_state.get("generated_master") or
#         (BytesIO(uploaded_master.read()) if uploaded_master else None)
#     )
#     stream_bytes = BytesIO(uploaded_stream.read()) if uploaded_stream else None

#     if master_bytes and stream_bytes:
#         master_bytes.seek(0)
#         stream_bytes.seek(0)

#         result, missing = populate_syscad_inputs(master_bytes, stream_bytes)

#         if missing:
#             st.warning(f"⚠️ Missing streamtable data for: {', '.join(missing)}")

#         st.success("✅ SysCAD inputs successfully populated into the master sheet.")

#         result.seek(0)
#         timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
#         category_name = "SysCAD"
#         populated_filename = f"Master_DataSheet_{category_name}Populated_{timestamp}.xlsx"

#         st.download_button(
#             label="📥 Download Populated Master Sheet",
#             data=result,
#             file_name=populated_filename,
#             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#         )
#     else:
#         st.error("Please either generate or upload a master sheet, and upload a streamtable file.")
