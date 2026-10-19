from datetime import date, datetime
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Area-1 C&I Plant Hub",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .stApp { background-color: #0b0f19; color: #e2e8f0; }
    .block-container { padding-top: 1.5rem; padding-bottom: 2.5rem; }
    div[data-testid="stMetric"] { background: #151c2c; border: 1px solid #243048; border-radius: 8px; padding: 12px 16px; }
    div[data-testid="stMetricLabel"] { font-size: 0.8rem; color: #94a3b8; font-weight: 600; text-transform: uppercase; }
    div[data-testid="stMetricValue"] { font-size: 1.6rem; font-weight: 700; color: #38bdf8; }
    .plant-header { font-size: 1.6rem; font-weight: 700; color: #f8fafc; margin-bottom: 4px; }
    .plant-sub { color: #94a3b8; font-size: 0.88rem; margin-bottom: 20px; }
    </style>
""",
    unsafe_allow_html=True,
)

PLANT_MODULES = {
    "📝 PR and PO Tracker": {
        "file": "PR and PO.xlsx",
        "desc": "Procurement indents, LTE/Open status, PO awards, and delivery due aging.",
        "type": "pr_po",
    },
    "⚡ PMPS Schedule": {
        "file": "PMPS.xlsx",
        "desc": "Preventive maintenance, calibration frequency, and overdue instrumentation tags.",
        "type": "pmps",
    },
    "📦 Spares Inventory": {
        "file": "Spares.xlsx",
        "desc": "Material codes, central store stock, installed units, and card allocations.",
        "type": "spares",
    },
    "🌐 IP Address Register": {
        "file": "IP address.xlsx",
        "desc": "DCS server nodes, operator stations, subnets, and gateway mappings.",
        "type": "ip",
    },
    "👥 C-Off & Attendance": {
        "file": "Coff Status.xlsx",
        "desc": "Compensatory off balances, extra duty days, and monthly attendance tracking.",
        "type": "coff",
    },
}

st.sidebar.markdown("## 🏭 Plant Operations Hub")
st.sidebar.caption("Area-1 Instrumentation & Control")
active_module = st.sidebar.radio(
    "Select Register:", list(PLANT_MODULES.keys())
)
selected_config = PLANT_MODULES[active_module]
current_file = selected_config["file"]

if not os.path.exists(current_file):
  st.error(f"File `{current_file}` not found.")
  st.stop()


@st.cache_data
def get_sheets(filename):
  return pd.ExcelFile(filename).sheet_names


sheet_names = get_sheets(current_file)
selected_sheet = st.sidebar.selectbox("Select Sheet / Period:", sheet_names)
st.sidebar.divider()
st.sidebar.markdown(f"**Workbook:** `{current_file}`")


@st.cache_data
def load_clean_data(filename, sheet):
  raw = pd.read_excel(filename, sheet_name=sheet)

  # Auto-detect header row
  header_idx = None
  for idx, row in raw.head(10).iterrows():
    row_str = " ".join([str(v).upper() for v in row.values if pd.notna(v)])
    if any(
        k in row_str
        for k in [
            "TAG NO",
            "DESCRIPTION",
            "INDENT",
            "NODE NAME",
            "EMPLOYEE",
            "MATERIAL CODE",
            "S.NO",
        ]
    ):
      header_idx = idx
      break

  if header_idx is not None:
    raw.columns = raw.iloc[header_idx]
    raw = raw.iloc[header_idx + 1 :].reset_index(drop=True)

  raw = raw.dropna(how="all", axis=0).dropna(how="all", axis=1)

  # Deduplicate and rename blank columns
  clean_cols = []
  seen = {}
  for i, c in enumerate(raw.columns):
    name = (
        f"Col_{i+1}"
        if pd.isna(c) or str(c).strip() == "" or "UNNAMED:" in str(c).upper()
        else str(c).strip()
    )
    if name in seen:
      seen[name] += 1
      name = f"{name}_{seen[name]}"
    else:
      seen[name] = 0
    clean_cols.append(name)
  raw.columns = clean_cols
  return raw.reset_index(drop=True)


df = load_clean_data(current_file, selected_sheet)

st.markdown(
    f'<div class="plant-header">{active_module} — {selected_sheet}</div>',
    unsafe_allow_html=True,
)
st.markdown(
    f'<div class="plant-sub">{selected_config["desc"]}</div>',
    unsafe_allow_html=True,
)

# --- MODULE CUSTOM LOGIC ---
if selected_config["type"] == "pr_po":
  due_col = [c for c in df.columns if "DEL. DUE" in str(c).upper()]
  po_col = [c for c in df.columns if "PO NO" in str(c).upper()]
  est_col = [c for c in df.columns if "EST VALUE" in str(c).upper()]

  if due_col:
    df[due_col[0]] = pd.to_datetime(df[due_col[0]], errors="coerce")
    today = pd.to_datetime(date.today())
    overdue = (df[due_col[0]] < today) & df[due_col[0]].notna()
    overdue_count = int(overdue.sum())
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Items Listed", len(df))
    c2.metric(
        "Overdue Deliveries",
        overdue_count,
        delta=f"-{overdue_count} Overdue" if overdue_count > 0 else "Clear",
        delta_color="inverse",
    )
    if est_col:
      val_clean = (
          df[est_col[0]]
          .astype(str)
          .str.replace("₹", "")
          .str.replace(",", "")
          .str.strip()
      )
      total_est = pd.to_numeric(val_clean, errors="coerce").sum()
      c3.metric("Requisition Total (INR)", f"₹ {total_est:,.0f}")
    if overdue_count > 0 and st.checkbox(
        "🚨 Filter: Show Overdue Orders Only", value=False
    ):
      df = df[overdue]
  else:
    c1, c2 = st.columns(2)
    c1.metric("Total Items", len(df))
    if po_col:
      c2.metric("POs Generated", df[po_col[0]].notna().sum())

elif selected_config["type"] == "pmps":
  due_col = [c for c in df.columns if "NEXT PMPS DUE" in str(c).upper()]
  crit_col = [c for c in df.columns if "CRIT" in str(c).upper()]
  c1, c2, c3 = st.columns(3)
  c1.metric("Instrumentation Tags", len(df))
  if crit_col:
    c2.metric("Essential (Critical-E) Tags", (df[crit_col[0]] == "E").sum())
  if due_col:
    c3.metric(
        "Tags with Scheduled Due Dates", df[due_col[0]].dropna().nunique()
    )

elif selected_config["type"] == "spares":
  stock_col = [c for c in df.columns if "STOCK" in str(c).upper()]
  inst_col = [c for c in df.columns if "INSTL" in str(c).upper()]
  c1, c2, c3 = st.columns(3)
  c1.metric("Line Items", len(df))
  if stock_col:
    stock_num = pd.to_numeric(df[stock_col[0]], errors="coerce").fillna(0)
    c2.metric("Total Physical Stock", f"{int(stock_num.sum()):,}")
    c3.metric(
        "Zero Stock Items",
        int((stock_num == 0).sum()),
        delta="-Restock" if (stock_num == 0).sum() > 0 else "Optimal",
        delta_color="inverse",
    )

elif selected_config["type"] == "ip":
  node_col = [c for c in df.columns if "NODE NAME" in str(c).upper()]
  loc_col = [c for c in df.columns if "LOCATION" in str(c).upper()]
  c1, c2 = st.columns(2)
  c1.metric("Network Nodes", len(df))
  if loc_col:
    c2.metric("Locations / Control Rooms", df[loc_col[0]].nunique())

else:  # coff
  c1, c2 = st.columns(2)
  c1.metric("Employee Records", len(df))
  extra_col = [c for c in df.columns if "EXTRA" in str(c).upper()]
  if extra_col:
    c2.metric(
        "Total Extra Days",
        int(pd.to_numeric(df[extra_col[0]], errors="coerce").fillna(0).sum()),
    )

st.divider()

# --- EDITABLE GRID & SAVE ---
st.subheader(f"Data Matrix — {selected_sheet}")
search_txt = st.text_input(
    "Search Data",
    placeholder="Search by tag, node, material code, employee, or description...",
    label_visibility="collapsed",
)
filtered_df = df.copy()
if search_txt:
  mask = filtered_df.astype(str).apply(
      lambda r: r.str.contains(search_txt, case=False).any(), axis=1
  )
  filtered_df = filtered_df[mask]

edited_df = st.data_editor(
    filtered_df, num_rows="dynamic", use_container_width=True, hide_index=True
)

st.markdown("<br>", unsafe_allow_html=True)
btn1, btn2 = st.columns([2, 1])
with btn1:
  if st.button(
      f"💾 Save Changes to '{selected_sheet}' in {current_file}",
      type="primary",
      use_container_width=True,
  ):
    with pd.ExcelWriter(
        current_file, engine="openpyxl", mode="a", if_sheet_exists="replace"
    ) as writer:
      edited_df.to_excel(writer, sheet_name=selected_sheet, index=False)
    st.cache_data.clear()
    st.success(f"Sheet '{selected_sheet}' updated successfully!")
    st.rerun()

with btn2:
  st.download_button(
      label=f"Export {selected_sheet} (CSV)",
      data=edited_df.to_csv(index=False),
      file_name=f"{selected_sheet}.csv",
      mime="text/csv",
      use_container_width=True,
  )
