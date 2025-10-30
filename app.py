import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="MasterCmd Parameter Assignment Configurator", layout="wide")

st.title("⚙️ MasterCmd Parameter Assignment Configurator")

# --- Step 1: User Inputs ---
st.sidebar.header("Configuration Settings")

num_blocks = st.sidebar.number_input("Blocks per Device", min_value=1, value=10, step=1)
nodes_sequence = st.sidebar.text_area(
    "Enter Node Numbers (comma-separated)",
    "26,27,28,29,30,31,32,33,34,35",
)
node_numbers = [int(n.strip()) for n in nodes_sequence.split(",") if n.strip().isdigit()]
num_devices = len(node_numbers)

entered_device_count = st.sidebar.number_input("Number of Devices (auto-calculated)", value=num_devices, step=1)
if entered_device_count != num_devices:
    st.sidebar.warning(f"⚠️ Entered number ({entered_device_count}) differs from detected nodes ({num_devices}).")

st.sidebar.write("")

# --- Step 2: Block Parameter Table Input ---
st.subheader("🧱 Block Parameter Configuration")

st.markdown("""
Enter the configuration values for each block.
- Leave **Enable** blank → treated as 0  
- **Enable** can be 0, 1, or 2  
- **Func** default = 4  
""")

default_data = {
    "Block No.": list(range(1, num_blocks + 1)),
    "Enable": [1] * num_blocks,
    "Func": [4] * num_blocks,
    "DevAddress": [100 + i for i in range(num_blocks)],
    "Count": [1] * num_blocks,
}
block_df = pd.DataFrame(default_data)
edited_df = st.data_editor(block_df, num_rows="dynamic", key="block_table")

# --- Step 3: Function Configuration ---
st.subheader("🧩 Function Configuration")

func_ids = sorted(edited_df["Func"].dropna().unique())
func_config = {}
cols = st.columns(len(func_ids))
for i, func_id in enumerate(func_ids):
    with cols[i]:
        st.markdown(f"**Function Code {func_id}**")
        init_addr = st.number_input(f"Initial Internal Address", min_value=0, value=0, key=f"init_{func_id}")
        offset = st.number_input(f"New Device Offset", min_value=0, value=10, key=f"offset_{func_id}")
        func_config[func_id] = {"initial": init_addr, "offset": offset}

# --- Step 4: Preview or Generate ---
col1, col2 = st.columns([1, 1])
generate_clicked = col1.button("💾 Generate Excel")
preview_clicked = col2.button("👁️ Preview Output")

def generate_excel():
    rows = []
    intaddress_track = {f: func_config[f]["initial"] for f in func_ids}

    for dev_idx, node in enumerate(node_numbers, start=1):
        for _, row in edited_df.iterrows():
            block_no = row["Block No."]
            enable = 0 if pd.isna(row["Enable"]) else int(row["Enable"])
            func = 4 if pd.isna(row["Func"]) else int(row["Func"])
            devaddr = int(row["DevAddress"])
            count = int(row["Count"])

            # Only calculate IntAddress for blocks with enable > 0
            if enable > 0:
                if dev_idx > 1 and row["Block No."] == 1:
                    # Apply new device offset only for first block of this func
                    intaddress_track[func] += func_config[func]["offset"]
                intaddr = intaddress_track[func]
                intaddress_track[func] += count
            else:
                intaddr = None

            # Add parameter rows (simplified pattern)
            params = [
                ("Enable", enable),
                ("Func", func),
                ("DevAddress", devaddr),
                ("Count", count),
                ("IntAddress", intaddr if intaddr is not None else ""),
                ("Node", node),
            ]
            for pname, val in params:
                rows.append({
                    "Device No.": dev_idx,
                    "Block No.": block_no,
                    "Node No.": node,
                    "Parameter": f"Cmd[{block_no}].{pname}",
                    "ConfigValue": val,
                })

    df = pd.DataFrame(rows)

    # Apply Excel formatting
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Port1"

    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    # Styling
    thin = Side(border_style="thick", color="000000")
    border = Border(top=thin, bottom=thin, left=thin, right=thin)
    fills = ["ADD8E6", "90EE90", "FFD580", "FFB6C1", "D3D3D3"]

    for idx, device in enumerate(df["Device No."].unique(), start=1):
        color = fills[(idx - 1) % len(fills)]
        dev_rows = df[df["Device No."] == device].index
        for i in dev_rows:
            for c in range(1, len(df.columns) + 1):
                ws.cell(row=i + 2, column=c).fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        ws.cell(row=dev_rows.min() + 2, column=1).border = border
        ws.cell(row=dev_rows.max() + 2, column=len(df.columns)).border = border

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return df, output

if preview_clicked:
    preview_df, _ = generate_excel()
    st.dataframe(preview_df)

if generate_clicked:
    df, output = generate_excel()
    st.success("✅ Excel file generated successfully!")
    st.download_button("⬇️ Download Excel", data=output, file_name="MasterCmd_Sequence_R01.xlsx")

# --- Save/Load Config ---
st.sidebar.subheader("⚙️ Save / Load Configuration")
if st.sidebar.button("💾 Save Current Configuration"):
    config_data = {
        "blocks": edited_df.to_dict(),
        "func_config": func_config,
        "nodes": node_numbers,
    }
    st.sidebar.download_button(
        label="⬇️ Download Config JSON",
        data=pd.Series(config_data).to_json(),
        file_name="config.json",
    )

uploaded_config = st.sidebar.file_uploader("Upload Config JSON", type=["json"])
if uploaded_config:
    import json
    loaded = json.load(uploaded_config)
    st.session_state["block_table"] = pd.DataFrame(loaded["blocks"])
    st.sidebar.success("✅ Configuration loaded successfully!")
