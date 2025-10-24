import streamlit as st
import pandas as pd
import json
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, PatternFill
from pathlib import Path

# === Excel and Data Generation Logic ===
def build_dataframe(devices, blocks, rows_per_block, node_seq, func_rules, block_rules):
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

    rows = []
    global_block_idx = 0
    prev_intaddr = {}
    prev_count = {}

    for device_no in range(1, devices + 1):
        node_no = node_seq[device_no - 1]

        for block_no in range(1, blocks + 1):
            idx = global_block_idx
            rule = block_rules[block_no]

            enable = rule.get("Enable", 0)
            enable = 0 if enable in ("", None) else enable
            func = rule.get("Func", "4") if enable != 0 else ""
            devaddr = rule.get("DevAddress", "") if enable != 0 else ""
            count = rule.get("Count", "") if enable != 0 else ""

            func_id = str(func)
            func_conf = func_rules.get(func_id, {"start": 0, "offset": 10})

            for p in param_template:
                param = p.format(i=idx)
                base = param.split('.')[-1]
                cfg = ""

                if base == "Enable":
                    cfg = enable
                elif base == "Func":
                    cfg = func
                elif base == "DevAddress":
                    cfg = devaddr
                elif base == "Count":
                    cfg = count
                elif base == "Node":
                    cfg = node_no
                elif base == "IntAddress":
                    if enable == 0 or func == "":
                        cfg = ""
                    else:
                        if func_id not in prev_intaddr:
                            cfg = func_conf["start"]
                        else:
                            cfg = prev_intaddr[func_id] + prev_count.get(func_id, 0)
                            if device_no > 1:
                                cfg += func_conf["offset"]
                        prev_intaddr[func_id] = cfg
                        prev_count[func_id] = int(count) if count != "" else 0

                rows.append({
                    "Device No.": device_no,
                    "Block No.": block_no,
                    "Node No.": node_no,
                    "Parameter": param,
                    "ConfigValue": cfg
                })

            global_block_idx += 1

    return pd.DataFrame(rows)


def generate_excel(df, output_name, devices, blocks, rows_per_block):
    output_path = Path(output_name)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Port1", index=False)
        pd.DataFrame(columns=df.columns).to_excel(writer, sheet_name="Port2", index=False)

    # === Formatting ===
    wb = load_workbook(output_path)
    ws = wb["Port1"]

    thick = Side(border_style="thick", color="000000")
    fill_colors = ["FFF2CC", "D9EAD3", "FCE5CD", "D0E0E3", "EAD1DC", "F4CCCC"]

    start_row = 2
    row_idx = start_row

    for device_no in range(1, devices + 1):
        for block_no in range(1, blocks + 1):
            color = fill_colors[(block_no - 1) % len(fill_colors)]
            fill = PatternFill(fill_type="solid", fgColor=color)

            block_start = row_idx
            block_end = block_start + rows_per_block - 1

            for r in range(block_start, block_end + 1):
                for c in range(1, 6):
                    ws.cell(row=r, column=c).fill = fill

            for col in range(1, 6):
                ws.cell(row=block_start, column=col).border = Border(top=thick)
                ws.cell(row=block_end, column=col).border = Border(bottom=thick)
            for rown in range(block_start, block_end + 1):
                ws.cell(row=rown, column=1).border = Border(left=thick)
                ws.cell(row=rown, column=5).border = Border(right=thick)

            row_idx = block_end + 1

    wb.save(output_path)
    return output_path


# === Streamlit UI ===
st.set_page_config(page_title="MasterCmd Parameter Assignment Configurator", layout="wide")
st.title("⚙️ MasterCmd Parameter Assignment Configurator")

st.markdown("""
Use this tool to build and export **MasterCmd configuration Excel files** with automatic IntAddress sequencing.
You can also save and reload configuration templates.
""")

# --- Load or Create Config ---
config_data = None
uploaded_config = st.file_uploader("📤 Upload a saved configuration (JSON)", type=["json"])
if uploaded_config:
    config_data = json.load(uploaded_config)
    st.success("Configuration loaded successfully!")

# --- Node sequence and device count ---
node_seq_input = st.text_input(
    "Enter Node Numbers (comma-separated):",
    config_data.get("node_seq_input", "26,27,28,29,30,31,32,33,34,35") if config_data else "26,27,28,29,30,31,32,33,34,35"
)
nodes = [int(x.strip()) for x in node_seq_input.split(",") if x.strip()]
auto_devices = len(nodes)

manual_devices = st.number_input("Number of Devices (optional)", min_value=1, value=config_data.get("manual_devices", auto_devices) if config_data else auto_devices)
if manual_devices != auto_devices:
    st.warning(f"⚠️ Entered device count ({manual_devices}) differs from node sequence ({auto_devices}).")
devices = auto_devices

# --- Other configuration ---
blocks = st.number_input("Blocks per Device", min_value=1, value=config_data.get("blocks", 10) if config_data else 10)
rows_per_block = st.number_input("Rows per Block", min_value=1, value=config_data.get("rows_per_block", 8) if config_data else 8)

st.markdown("### 🔢 Block Configuration (per device)")
st.write("Enter parameters per block. If `Enable` is blank or `0`, other fields can be left blank.")

block_rules = {}
cols = st.columns([1, 1, 1, 1, 1])
cols[0].write("**Block No.**")
cols[1].write("**Enable (0/1/2)**")
cols[2].write("**Func**")
cols[3].write("**DevAddress**")
cols[4].write("**Count**")

func_values = set()

for b in range(1, blocks + 1):
    prev = config_data["block_rules"].get(str(b), {}) if config_data else {}
    c = st.columns([1, 1, 1, 1, 1])
    c[0].write(f"{b}")
    enable = c[1].text_input(f"Enable_{b}", value=str(prev.get("Enable", "")), label_visibility="collapsed")
    enable = int(enable) if enable.isdigit() else 0
    func = c[2].text_input(f"Func_{b}", value=str(prev.get("Func", "4")), label_visibility="collapsed") if enable != 0 else ""
    devaddr = c[3].text_input(f"DevAddr_{b}", value=str(prev.get("DevAddress", "")), label_visibility="collapsed") if enable != 0 else ""
    count = c[4].number_input(f"Count_{b}", min_value=0, value=int(prev.get("Count", b)), label_visibility="collapsed") if enable != 0 else 0
    block_rules[b] = {"Enable": enable, "Func": func, "DevAddress": devaddr, "Count": count}
    if func not in ("", None):
        func_values.add(func)

# --- Function Configuration ---
st.markdown("### 🧩 Function Configuration")
st.write("Automatically detected from unique **Function Codes** entered above.")

cols = st.columns([1, 1, 1])
cols[0].write("**Function Code**")
cols[1].write("**Initial Internal Address**")
cols[2].write("**New Device Offset**")

func_rules = {}
for func_id in sorted(func_values):
    prev = config_data["func_rules"].get(func_id, {}) if config_data else {}
    cols = st.columns([1, 1, 1])
    cols[0].write(f"{func_id}")
    start = cols[1].number_input(f"Start_Int_{func_id}", value=int(prev.get("start", int(func_id) * 10)), label_visibility="collapsed")
    offset = cols[2].number_input(f"Offset_{func_id}", value=int(prev.get("offset", 10)), label_visibility="collapsed")
    func_rules[func_id] = {"start": start, "offset": offset}

output_name = st.text_input("Output Excel File Name", config_data.get("output_name", "MasterCmd_Sequence_Generated.xlsx") if config_data else "MasterCmd_Sequence_Generated.xlsx")

# --- Save Config Button ---
if st.button("💾 Save Configuration"):
    config_to_save = {
        "node_seq_input": node_seq_input,
        "manual_devices": manual_devices,
        "blocks": blocks,
        "rows_per_block": rows_per_block,
        "block_rules": block_rules,
        "func_rules": func_rules,
        "output_name": output_name
    }
    save_name = "MasterCmd_Config_Saved.json"
    with open(save_name, "w") as f:
        json.dump(config_to_save, f, indent=4)
    st.download_button("⬇️ Download Saved Config", data=json.dumps(config_to_save, indent=4),
                       file_name=save_name, mime="application/json")

# --- Preview Button ---
if st.button("👁 Preview Port1 Table"):
    with st.spinner("Generating preview..."):
        df_preview = build_dataframe(devices, blocks, rows_per_block, nodes, func_rules, block_rules)
        st.success("✅ Preview generated below")
        st.dataframe(df_preview, use_container_width=True, height=500)

# --- Generate Excel Button ---
if st.button("🚀 Generate Excel File"):
    with st.spinner("Generating Excel file..."):
        try:
            df = build_dataframe(devices, blocks, rows_per_block, nodes, func_rules, block_rules)
            output_path = generate_excel(df, output_name, devices, blocks, rows_per_block)
            with open(output_path, "rb") as f:
                st.success("✅ Excel file generated successfully!")
                st.download_button("⬇️ Download Excel File", data=f, file_name=output_name,
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"❌ Error: {e}")
