# MasterCmd Sequence Generator App
import streamlit as st
import pandas as pd
import json
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, PatternFill


# === Excel Generation Function ===
def generate_excel(devices, blocks, rows_per_block, node_seq, rules, preview=False):
    param_template = [
        "MCM.CONFIG.Port1MasterCmd[{i}].Enable",
        "MCM.CONFIG.Port1MasterCmd[{i}].IntAddress",
        "MCM.CONFIG.Port1MasterCmd[{i}].PollInt",
        "MCM.CONFIG.Port1MasterCmd[{i}].Count",
        "MCM.CONFIG.Port1MasterCmd[{i}].Swap",
        "MCM.CONFIG.Port1MasterCmd[{i}].Node",
        "MCM.CONFIG.Port1MasterCmd[{i}].Func",
        "MCM.CONFIG.Port1MasterCmd[{i}].DevAddress",
    ]

    enable_value = rules["enable"]
    count_map = rules["count_map"]
    devaddr_map = rules["devaddr_map"]
    func_map = rules["func_map"]
    func_init = rules["func_init"]
    func_offset = rules["func_offset"]

    rows = []
    global_block_idx = 0
    func_state = {f: func_init[f] for f in func_init}

    for device_no in range(1, devices + 1):
        node_no = node_seq[device_no - 1]
        for block_no in range(1, blocks + 1):
            func_val = func_map[block_no]
            func_intaddr = func_state[func_val]
            count_val = count_map[block_no]
            devaddr_val = devaddr_map[block_no]

            for p in param_template:
                param = p.format(i=global_block_idx)
                base = param.split('.')[-1]
                cfg = ""

                if base == "Enable":
                    cfg = enable_value
                elif base == "IntAddress":
                    cfg = func_intaddr
                elif base == "Count":
                    cfg = count_val
                elif base == "Node":
                    cfg = node_no
                elif base == "Func":
                    cfg = func_val
                elif base == "DevAddress":
                    cfg = devaddr_val

                rows.append({
                    "Device No.": device_no,
                    "Block No.": block_no,
                    "Node No.": node_no,
                    "Parameter": param,
                    "ConfigValue": cfg
                })

            func_state[func_val] = func_intaddr + count_val
            global_block_idx += 1

        for f in func_state:
            func_state[f] += func_offset[f]

    df = pd.DataFrame(rows)
    if preview:
        return df

    # === Excel Formatting ===
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Port1", index=False)
        pd.DataFrame(columns=df.columns).to_excel(writer, sheet_name="Port2", index=False)
    buf.seek(0)

    wb = load_workbook(buf)
    ws = wb["Port1"]
    thick = Side(border_style="thick", color="000000")
    block_colors = ["FFF2CC", "D9EAD3", "FCE5CD", "EAD1DC", "C9DAF8", "D0E0E3"]

    start_row = 2
    row_idx = start_row
    for device_no in range(1, devices + 1):
        for block_no in range(1, blocks + 1):
            block_start = row_idx
            block_end = block_start + rows_per_block - 1
            fill_color = block_colors[(block_no - 1) % len(block_colors)]
            fill = PatternFill(fill_type="solid", fgColor=fill_color)

            for r in range(block_start, block_end + 1):
                for c in range(1, 6):
                    ws.cell(row=r, column=c).fill = fill
                    if r == block_start:
                        ws.cell(row=r, column=c).border = Border(top=thick)
                    if r == block_end:
                        ws.cell(row=r, column=c).border = Border(bottom=thick)
                    if c == 1:
                        ws.cell(row=r, column=c).border = Border(left=thick)
                    if c == 5:
                        ws.cell(row=r, column=c).border = Border(right=thick)
            row_idx += rows_per_block

    out_buf = BytesIO()
    wb.save(out_buf)
    out_buf.seek(0)
    return out_buf


# === Streamlit Interface ===
st.title("⚙️ MasterCmd Sequence Generator (Advanced Web Version)")
st.caption("Developed by Mohammed Elmehy")

# --- Sidebar Save/Load ---
st.sidebar.header("💾 Configuration Management")
if "config_data" not in st.session_state:
    st.session_state["config_data"] = {}

uploaded_cfg = st.sidebar.file_uploader("Load Config (.json)", type=["json"])
if uploaded_cfg:
    st.session_state["config_data"] = json.load(uploaded_cfg)
    st.sidebar.success("✅ Config loaded successfully!")

# --- Main Inputs ---
devices = st.number_input("Number of Devices", min_value=1, value=st.session_state["config_data"].get("devices", 26))
blocks = st.number_input("Blocks per Device", min_value=1, value=st.session_state["config_data"].get("blocks", 6))
rows_per_block = st.number_input("Rows per Block", min_value=1, value=st.session_state["config_data"].get("rows_per_block", 8))
node_str = st.text_input(
    "Node Numbers (comma-separated)",
    st.session_state["config_data"].get(
        "node_str",
        "26,27,28,29,30,31,32,33,34,35,61,62,63,64,65,66,67,68,69,70,71,72,73,74,75,76"
    )
)
node_seq = [int(x.strip()) for x in node_str.split(",")]

st.subheader("General Parameters")
enable_value = st.number_input(".Enable Value", value=st.session_state["config_data"].get("enable", 1))

st.markdown("### Block Configuration")
count_map, devaddr_map, func_map = {}, {}, {}
func_values = []

cfg_blocks = st.session_state["config_data"].get("blocks_data", {})

for b in range(1, blocks + 1):
    st.markdown(f"**Block {b}**")
    c1, c2, c3 = st.columns(3)
    with c1:
        count_map[b] = int(st.number_input(
            f"Count (Block {b})", value=cfg_blocks.get(str(b), {}).get("count", 1 if b != 2 else 5), key=f"count_{b}"
        ))
    with c2:
        devaddr_map[b] = int(st.number_input(
            f"DevAddress (Block {b})", value=cfg_blocks.get(str(b), {}).get("devaddr", {1:3, 2:101, 3:116, 4:142}.get(b, 0)), key=f"dev_{b}"
        ))
    with c3:
        func_map[b] = int(st.number_input(
            f"Func (Block {b})", value=cfg_blocks.get(str(b), {}).get("func", 3), key=f"func_{b}"
        ))
        func_values.append(func_map[b])

unique_funcs = sorted(set(func_values))
st.markdown("### Func Configuration (Initial .IntAddress & Offset per Device)")

cfg_funcs = st.session_state["config_data"].get("func_data", {})
func_init, func_offset = {}, {}
for f in unique_funcs:
    c1, c2 = st.columns(2)
    with c1:
        func_init[f] = int(st.number_input(
            f"Initial .IntAddress for Func {f}", value=cfg_funcs.get(str(f), {}).get("init", 0), key=f"init_{f}"
        ))
    with c2:
        func_offset[f] = int(st.number_input(
            f"Offset per New Device for Func {f}", value=cfg_funcs.get(str(f), {}).get("offset", 10), key=f"offset_{f}"
        ))

rules = {
    "enable": enable_value,
    "count_map": count_map,
    "devaddr_map": devaddr_map,
    "func_map": func_map,
    "func_init": func_init,
    "func_offset": func_offset,
}

# --- Save Configuration to JSON ---
config_to_save = {
    "devices": devices,
    "blocks": blocks,
    "rows_per_block": rows_per_block,
    "node_str": node_str,
    "enable": enable_value,
    "blocks_data": {
        str(b): {"count": count_map[b], "devaddr": devaddr_map[b], "func": func_map[b]}
        for b in range(1, blocks + 1)
    },
    "func_data": {
        str(f): {"init": func_init[f], "offset": func_offset[f]} for f in unique_funcs
    }
}

st.sidebar.download_button(
    "💾 Save Current Config",
    data=json.dumps(config_to_save, indent=4),
    file_name="MasterCmd_Config.json",
    mime="application/json",
)

# === Preview Button ===
if st.button("Preview Table"):
    if len(node_seq) != devices:
        st.error("❌ Number of node numbers must match number of devices.")
    else:
        df_preview = generate_excel(devices, blocks, rows_per_block, node_seq, rules, preview=True)
        st.success("✅ Preview generated successfully!")
        st.dataframe(df_preview.head(100))

# === Generate Excel Button ===
if st.button("Generate Excel File"):
    if len(node_seq) != devices:
        st.error("❌ Number of node numbers must match number of devices.")
    else:
        buf = generate_excel(devices, blocks, rows_per_block, node_seq, rules)
        st.success("✅ Excel generated successfully!")
        st.download_button(
            label="📥 Download Excel",
            data=buf,
            file_name="MasterCmd_Sequence.Advanced.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

