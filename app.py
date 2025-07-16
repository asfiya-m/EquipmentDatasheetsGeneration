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
from io import BytesIO
import pandas as pd
import streamlit as st
from datetime import datetime

from automation_test1 import generate_master_datasheet
from populate_equipment_names import populate_equipment_names
from populate_parameters import populate_parameters

st.title("📄 Master Equipment Datasheet Automation Tool")

st.markdown("""
This tool helps you:
1. Generate a clean, categorized master datasheet from the Excel datasheets workbook.
2. Populate Equipment names.
3. Populate SysCAD parameter values.
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
    output_stream.seek(0)
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

use_generated_step2 = False
if "generated_master" in st.session_state:
    use_generated_step2 = st.radio(
        "Choose master sheet to populate:",
        ["Use the one generated in Step 1", "Upload a different master sheet"],
        key="step2_radio"
    ) == "Use the one generated in Step 1"

if use_generated_step2:
    master_bytes = st.session_state["generated_master"]
else:
    uploaded_master = st.file_uploader("Upload the master sheet", type=["xlsx"], key="master2")
    if uploaded_master:
        master_bytes = BytesIO(uploaded_master.read())
    else:
        master_bytes = None

uploaded_stream = st.file_uploader("Upload the detailed streamtable", type=["xlsx"], key="stream2")

if master_bytes and uploaded_stream and st.button("Populate Equipment Names"):
    stream_content = uploaded_stream.read()  # read bytes here
    stream_bytes = BytesIO(stream_content)

    result, filename, skipped = populate_equipment_names(master_bytes, stream_bytes)

    # Save outputs for Step 3
    result.seek(0)
    st.session_state["master_with_equipment"] = result
    st.session_state["uploaded_stream_content"] = stream_content

    if skipped:
        st.warning(f"⚠️ Some equipment were not matched to any sheet:\n{', '.join(skipped)}")

    st.success("✅ Equipment names populated successfully.")

    st.download_button(
        label="📥 Download Populated Master Sheet",
        data=result,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ------------------------
# Step 3: Populate Parameter values
# ------------------------
st.header("Step 3: Populate Parameters")
st.markdown("""
**What happens in this step?**
- Uses the populated master sheet (with equipment names from Step 2).
- Reads equipment & stream tags from your detailed streamtable.
- Looks up stream data in **Stream Table V**.
- Maps & writes parameters into the master datasheet under **SysCAD Inputs** section.
- Applies rules: sum, average, unit conversion as specified.
""")

use_generated_step3 = False
if "master_with_equipment" in st.session_state:
    use_generated_step3 = st.radio(
        "Choose master sheet to populate:",
        ["Use the one generated in Step 2", "Upload a different master sheet"],
        key="step3_radio"
    ) == "Use the one generated in Step 2"

if use_generated_step3:
    master_bytes = st.session_state["master_with_equipment"]
else:
    uploaded_master = st.file_uploader("Upload the master sheet", type=["xlsx"], key="master3")
    if uploaded_master:
        master_bytes = BytesIO(uploaded_master.read())
    else:
        master_bytes = None

if "uploaded_stream_content" in st.session_state:
    stream_bytes = BytesIO(st.session_state["uploaded_stream_content"])
else:
    stream_bytes = None

if master_bytes and stream_bytes and st.button("Populate Parameters"):
    result, filename, skipped = populate_parameters(master_bytes, stream_bytes)

    if skipped:
        st.warning("⚠️ Some equipment or streams could not be populated.")
        st.text_area("Skipped Items", "\n".join(skipped), height=200)

        skipped_df = pd.DataFrame(skipped, columns=["Skipped"])
        skipped_csv = skipped_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="📥 Download Skipped Items List",
            data=skipped_csv,
            file_name="skipped_parameters.csv",
            mime="text/csv"
        )

    st.success("✅ Parameters populated successfully into master sheet.")

    st.download_button(
        label="📥 Download Populated Master Sheet",
        data=result,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
