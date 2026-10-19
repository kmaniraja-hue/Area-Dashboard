from datetime import date, datetime
import os
import pandas as pd
import streamlit as st

# 1. Page Configuration
st.set_page_config(
    page_title="Area Dashboard — Plant Asset & Operations Portal",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

# 2. Industrial Theme Styling
st.markdown(
    """
    <style>
    .stApp {
        background-color: #0b0f19;
        color: #e2e8f0;
    }
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2.5rem;
    }
    div[data-testid="stMetric"] {
        background: #151c2c;
        border: 1px solid #243048;
        border-radius: 8px;
        padding: 12px 16px;
        box-shadow: 0 4px 10px rgba(0,0,0,0.3);
    }
    div[data-testid="stMetricLabel"] {
        font-size: 0.8rem;
        color: #94a3b8;
        font-weight: 600;
        text-transform: uppercase;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.6rem;
        font-weight: 700;
        color: #38bdf8;
    }
    .plant-header {
        font-size: 1.6rem;
        font-weight: 700;
        color: #f8fafc;
        margin-bottom: 4px;
    }
    .plant-sub {
        color: #94a3b8;
        font-size: 0.88rem;
        margin-bottom: 20px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# 3. Exact File Mapping Matching Your GitHub Upload
PLANT_MODULES = {
    "🌐 IP Address Register": {
        "file": "IP address.xlsx",
        "desc": (
            "DCS station IPs, network switch allocations, subnet & gateway"
            " records."
        ),
        "type": "ip",
    },
    "🚰 Coff Status (Cut-Off Valves)": {
        "file": "Coff Status.xlsx",
        "desc": (
            "Cut-off valves, line sizes, service area allocations, and"
            " maintenance status."
        ),
        "type": "coff",
    },
    "⚡ PMPS (Power Modules & Supplies)": {
        "file": "PMPS.xlsx",
        "desc": (
            "PMIO & DCS power supply units, redundancy checks, and health logs."
        ),
        "type": "pmps",
    },
    "📝 PR and PO Tracker": {
        "file": "PR and PO.xlsx",
        "desc": (
            "Purchase requisitions, purchase orders, delivery due dates, and"
            " overdue alarms."
        ),
        "type": "pr_po",
    },
    "📦 Spares Master Inventory": {
        "file": "Spares.xlsx",
        "desc": (
            "Control room spares allocation, cards, processors, and station"
            " stock levels."
        ),
        "type": "spares",
    },
}

# 4. Sidebar Navigation
st.sidebar.markdown("## 🏭 Plant Asset Hub")
st.sidebar.caption("Area Dashboard — C&I Operations")

active_module = st.sidebar.radio(
    "Select Active Register:", list(PLANT_MODULES.keys())
)

selected_config = PLANT_MODULES[active_module]
current_file = selected_config["file"]

st.sidebar.divider()
st.sidebar.markdown(f"**Loaded Workbook:** `{current_file}`")


# 5. Generic Excel Loader
@st.cache_data
def load_excel_sheet(filename):
  if not os.path.exists(filename):
    return None
  try:
    raw = pd.read_excel(filename)
    # Search top 10 rows to detect header offset
    for idx, row in raw.head(10).iterrows():
      row_str = " ".join([str(v).upper() for v in row.values])
      if any(
          k in row_str
          for k in [
              "IP",
              "TAG",
              "S.NO.",
              "PO NO",
              "PR NO",
              "ITEM",
              "MAKE",
              "STATUS",
              "QTY",
              "STATION",
          ]
      ):
        raw.columns = raw.iloc[idx]
        raw = raw.iloc[idx + 1 :].reset_index(drop=True)
        break
    return raw.dropna(how="all")
  except Exception as e:
    st.error(f"Error loading {filename}: {e}")
    return None


df = load_excel_sheet(current_file)

if df is None:
  st.warning(f"File `{current_file}` not found in repository.")
  st.stop()

# 6. Page Header
st.markdown(
    f'<div class="plant-header">{active_module}</div>', unsafe_allow_html=True
)
st.markdown(
    f'<div class="plant-sub">{selected_config["desc"]}</div>',
    unsafe_allow_html=True,
)

# ------------------------------------------------------------------------------
# MODULE 1: IP ADDRESS REGISTER
# ------------------------------------------------------------------------------
if selected_config["type"] == "ip":
  loc_col = [
      c
      for c in df.columns
      if "LOC" in str(c).upper() or "ROOM" in str(c).upper()
  ]
  dev_col = [
      c
      for c in df.columns
      if "TYPE" in str(c).upper() or "DEVICE" in str(c).upper()
  ]

  i1, i2, i3 = st.columns(3)
  i1.metric("Assigned IP Records", len(df))
  if loc_col:
    i2.metric("Locations / Rooms", df[loc_col[0]].nunique())
  if dev_col:
    i3.metric("Hardware Device Types", df[dev_col[0]].nunique())

  st.markdown("<br>", unsafe_allow_html=True)
  if loc_col:
    st.subheader("📊 IP Distribution by Location / Substation")
    loc_chart = df[loc_col[0]].value_counts().reset_index()
    loc_chart.columns = ["Location", "IP Count"]
    st.bar_chart(
        loc_chart.set_index("Location"), color="#38bdf8", use_container_width=True
    )

# ------------------------------------------------------------------------------
# MODULE 2: COFF STATUS (CUT-OFF VALVES)
# ------------------------------------------------------------------------------
elif selected_config["type"] == "coff":
  area_col = [
      c
      for c in df.columns
      if "AREA" in str(c).upper() or "PHASE" in str(c).upper()
  ]
  stat_col = [
      c
      for c in df.columns
      if "STATUS" in str(c).upper() or "HEALTH" in str(c).upper()
  ]

  c1, c2, c3 = st.columns(3)
  c1.metric("Cut-Off Valve Population", len(df))
  if area_col:
    c2.metric("Plant Areas Covered", df[area_col[0]].nunique())
  if stat_col:
    needs_action = int(
        df[stat_col[0]]
        .astype(str)
        .str.upper()
        .str.contains("DUE|FAIL|LEAK|FAULT")
        .sum()
    )
    c3.metric(
        "Maintenance Alerts",
        needs_action,
        delta="-Attention" if needs_action > 0 else "Normal",
        delta_color="inverse",
    )

  st.markdown("<br>", unsafe_allow_html=True)
  if area_col:
    st.subheader("📊 Valve Population Across Plant Areas")
    v_chart = df[area_col[0]].value_counts().reset_index()
    v_chart.columns = ["Area", "Valves"]
    st.bar_chart(
        v_chart.set_index("Area"), color="#06b6d4", use_container_width=True
    )

# ------------------------------------------------------------------------------
# MODULE 3: PMPS (POWER MODULES & SUPPLIES)
# ------------------------------------------------------------------------------
elif selected_config["type"] == "pmps":
  stn_col = [
      c
      for c in df.columns
      if "STATION" in str(c).upper() or "ROOM" in str(c).upper()
  ]
  make_col = [c for c in df.columns if "MAKE" in str(c).upper()]
  health_col = [c for c in df.columns if "HEALTH" in str(c).upper()]

  k1, k2, k3 = st.columns(3)
  k1.metric("PMPS Modules Tracked", len(df))
  if make_col:
    k2.metric("Manufacturers / Makes", df[make_col[0]].nunique())
  if health_col:
    aging = int(
        df[health_col[0]]
        .astype(str)
        .str.upper()
        .str.contains("AGING|FAULT|DEFECT")
        .sum()
    )
    k3.metric(
        "Units Needing Attention",
        aging,
        delta="-Action Needed" if aging > 0 else "Normal",
        delta_color="inverse",
    )

  st.markdown("<br>", unsafe_allow_html=True)
  if stn_col:
    st.subheader("📊 Power Supplies Distribution by Station")
    stn_counts = df[stn_col[0]].value_counts().reset_index()
    stn_counts.columns = ["Station", "Units"]
    st.bar_chart(
        stn_counts.set_index("Station"), color="#6366f1", use_container_width=True
    )

# ------------------------------------------------------------------------------
# MODULE 4: PR AND PO TRACKER (WITH OVERDUE ENGINE)
# ------------------------------------------------------------------------------
elif selected_config["type"] == "pr_po":
  due_col = [c for c in df.columns if "DUE" in str(c).upper()]
  status_col = [
      c
      for c in df.columns
      if "STATUS" in str(c).upper() or "STAGE" in str(c).upper()
  ]

  if due_col:
    dcol = due_col[0]
    df[dcol] = pd.to_datetime(df[dcol], errors="coerce")
    today = pd.to_datetime(date.today())

    def check_status(row):
      stat = str(row.get(status_col[0], "")).upper() if status_col else ""
      if "DELIVER" in stat or "RECEIVED" in stat:
        return "Delivered"
      if pd.isna(row[dcol]):
        return "Date Missing"
      if row
