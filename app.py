import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, PatternFill
from pathlib import Path


# === Excel Generation Logic ===
def generate_excel(devices, blocks, rows_per_block, node_seq, func_rules, block_rules, output_name):
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
            enable = rule["Enable"]
            func = rule["Func"] if enable != 0 else ""
            devaddr = rule["DevAddress"] if enable != 0 else ""
            count = rule["Count"] if enable != 0 else ""

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

    df = pd.DataFrame(rows)

    # === Save to Excel ===
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
st.set_page_config(page_title="MasterCmd Sequence Generator", layout="wide")

st.title("⚙️ MasterCmd Excel Sequence Generator")

st.markdown("""
Generate **Master Command configuration Excel** files with flexible block-based parameters and automatic IntAddress sequencing.
""")

# --- Node sequence and device count ---
node_seq_input = st.text_input("Enter Node Numbers (comma-separated):",
                               "26,27,28,29,30,31,32,33,34,35,61,62,63,64,65,66,67,68,69,70,71,72,73,74,75,76")
nodes = [int(x.strip()) for x in node_seq_input.split(",") if x.strip()]
auto_devices = len(nodes)

manual_devices = st.number_input("Number of Devices (optional)", min_value=1, value=auto_devices)
if manual_devices != auto_devices:
    st.warning(f"⚠️ Entered device count ({manual_devices}) differs from node sequence ({auto_devices}).")

devices = auto_devices

# --- Other configuration ---
blocks = st.number_input("Blocks per Device", min_value=1, value=6)
rows_per_block = st.number_input("Rows per Block", min_value=1, value=8)

st.markdown("### 🔢 Block Configuration (per device)")
st.write("Enter parameters per block. If `Enable = 0`, other fields can be left blank.")

block_rules = {}
cols = st.columns([1, 1, 1, 1, 1])
cols[0].write("**Block No.**")
cols[1].write("**Enable (0/1/2)**")
cols[2].write("**Func**")
cols[3].write("**DevAddress**")
cols[4].write("**Count**")

func_values = set()

for b in range(1, blocks + 1):
    c = st.columns([1, 1, 1, 1, 1])
    c[0].write(f"{b}")
    enable = c[1].number_input(f"Enable_{b}", min_value=0, max_value=2, value=1, label_visibility="collapsed")
    func = c[2].text_input(f"Func_{b}", value=str(b), label_visibility="collapsed") if enable != 0 else ""
    devaddr = c[3].text_input(f"DevAddr_{b}", value=str(100 + b), label_visibility="collapsed") if enable != 0 else ""
    count = c[4].number_input(f"Count_{b}", min_value=0, value=b, label_visibility="collapsed") if enable != 0 else 0
    block_rules[b] = {"Enable": enable, "Func": func, "DevAddress": devaddr, "Count": count}
    if func not in ("", None):
        func_values.add(func)

# --- Func Configuration ---
st.markdown("### 🧩 Function Configuration")
st.write("Automatically recognized from unique Func IDs entered above.")

func_rules = {}
for func_id in sorted(func_values):
    cols = st.columns(3)
    cols[0].write(f"**Func ID {func_id}**")
    start = cols[1].number_input(f"Start_Int_{func_id}", value=int(func_id) * 10, label_visibility="collapsed")
    offset = cols[2].number_input(f"Offset_{func_id}", value=10, label_visibility="collapsed")
    func_rules[func_id] = {"start": start, "offset": offset}

output_name = st.text_input("Output Excel File Name", "MasterCmd_Sequence_Generated.xlsx")

# --- Generate Button ---
if st.button("🚀 Generate Excel File"):
    with st.spinner("Generating Excel file..."):
        try:
            output_path = generate_excel(devices, blocks, rows_per_block, nodes, func_rules, block_rules, output_name)
            with open(output_path, "rb") as f:
                st.success("✅ Excel file generated successfully!")
                st.download_button("⬇️ Download Excel File", data=f, file_name=output_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"❌ Error: {e}")
