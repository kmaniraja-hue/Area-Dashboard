import streamlit as st
import pandas as pd
from datetime import datetime, date
import os

# 1. Page Configuration
st.set_page_config(
    page_title="Area Dashboard — Plant Asset & Operations Portal",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
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
    unsafe_allow_html=True
)

# 3. File Mapping
PLANT_MODULES = {
    "🌐 IP Address Register": {
        "file": "IP address.xlsx",
        "desc": "DCS station IPs, network switch allocations, subnet & gateway records.",
        "type": "ip"
    },
    "🚰 Coff Status (Cut-Off Valves)": {
        "file": "Coff Status.xlsx",
        "desc": "Cut-off valves, line sizes, service area allocations, and maintenance status.",
        "type": "coff"
    },
    "⚡ PMPS (Power Modules & Supplies)": {
        "file": "PMPS.xlsx",
        "desc": "PMIO & DCS power supply units, redundancy checks, and health logs.",
        "type": "pmps"
    },
    "📝 PR and PO Tracker": {
        "file": "PR and PO.xlsx",
        "desc": "Purchase requisitions, purchase orders, delivery due dates, and overdue alarms.",
        "type": "pr_po"
    },
    "📦 Spares Master Inventory": {
        "file": "Spares.xlsx",
        "desc": "Control room spares allocation, cards, processors, and station stock levels.",
        "type": "spares"
    }
}

# 4. Sidebar Navigation
st.sidebar.markdown("## 🏭 Plant Asset Hub")
st.sidebar.caption("Area Dashboard — C&I Operations")

active_module = st.sidebar.radio("Select Active Register:", list(PLANT_MODULES.keys()))
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
        for idx, row in raw.head(10).iterrows():
            row_str = " ".join([str(v).upper() for v in row.values])
            if any(k in row_str for k in ["IP", "TAG", "S.NO.", "PO NO", "PR NO", "ITEM", "MAKE", "STATUS", "QTY", "STATION"]):
                raw.columns = raw.iloc[idx]
                raw = raw.iloc[idx + 1:].reset_index(drop=True)
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
st.markdown(f'<div class="plant-header">{active_module}</div>', unsafe_allow_html=True)
st.markdown(f'<div class="plant-sub">{selected_config["desc"]}</div>', unsafe_allow_html=True)

# Module 1: IP Register
if selected_config["type"] == "ip":
    loc_col = [c for c in df.columns if "LOC" in str(c).upper() or "ROOM" in str(c).upper()]
    dev_col = [c for c in df.columns if "TYPE" in str(c).upper() or "DEVICE" in str(c).upper()]

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
        st.bar_chart(loc_chart.set_index("Location"), color="#38bdf8", use_container_width=True)

# Module 2: Coff Status
elif selected_config["type"] == "coff":
    area_col = [c for c in df.columns if "AREA" in str(c).upper() or "PHASE" in str(c).upper()]
    stat_col = [c for c in df.columns if "STATUS" in str(c).upper() or "HEALTH" in str(c).upper()]

    c1, c2, c3 = st.columns(3)
    c1.metric("Cut-Off Valve Population", len(df))
    if area_col:
        c2.metric("Plant Areas Covered", df[area_col[0]].nunique())
    if stat_col:
        needs_action = int(df[stat_col[0]].astype(str).str.upper().str.contains("DUE|FAIL|LEAK|FAULT").sum())
        c3.metric("Maintenance Alerts", needs_action, delta="-Attention" if needs_action > 0 else "Normal", delta_color="inverse")

    st.markdown("<br>", unsafe_allow_html=True)
    if area_col:
        st.subheader("📊 Valve Population Across Plant Areas")
        v_chart = df[area_col[0]].value_counts().reset_index()
        v_chart.columns = ["Area", "Valves"]
        st.bar_chart(v_chart.set_index("Area"), color="#06b6d4", use_container_width=True)

# Module 3: PMPS
elif selected_config["type"] == "pmps":
    stn_col = [c for c in df.columns if "STATION" in str(c).upper() or "ROOM" in str(c).upper()]
    make_col = [c for c in df.columns if "MAKE" in str(c).upper()]
    health_col = [c for c in df.columns if "HEALTH" in str(c).upper()]

    k1, k2, k3 = st.columns(3)
    k1.metric("PMPS Modules Tracked", len(df))
    if make_col:
        k2.metric("Manufacturers / Makes", df[make_col[0]].nunique())
    if health_col:
        aging = int(df[health_col[0]].astype(str).str.upper().str.contains("AGING|FAULT|DEFECT").sum())
        k3.metric("Units Needing Attention", aging, delta="-Action Needed" if aging > 0 else "Normal", delta_color="inverse")

    st.markdown("<br>", unsafe_allow_html=True)
    if stn_col:
        st.subheader("📊 Power Supplies Distribution by Station")
        stn_counts = df[stn_col[0]].value_counts().reset_index()
        stn_counts.columns = ["Station", "Units"]
        st.bar_chart(stn_counts.set_index("Station"), color="#6366f1", use_container_width=True)

# Module 4: PR & PO Tracker
elif selected_config["type"] == "pr_po":
    due_col = [c for c in df.columns if "DUE" in str(c).upper()]
    status_col = [c for c in df.columns if "STATUS" in str(c).upper() or "STAGE" in str(c).upper()]

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
            if row[dcol] < today:
                return "OVERDUE"
            days_left = (row[dcol] - today).days
            if days_left <= 15:
                return "Due Soon (<15 Days)"
            return "On Schedule"

        df["Delivery Health"] = df.apply(check_status, axis=1)

        overdue_count = int((df["Delivery Health"] == "OVERDUE").sum())
        due_soon_count = int((df["Delivery Health"] == "Due Soon (<15 Days)").sum())
        delivered_count = int((df["Delivery Health"] == "Delivered").sum())

        p1, p2, p3, p4 = st.columns(4)
        p1.metric("Total Line Items", len(df))
        p2.metric("Critical Overdue", overdue_count, delta=f"-{overdue_count} Overdue" if overdue_count > 0 else "Clear", delta_color="inverse")
        p3.metric("Due in Next 15 Days", due_soon_count)
        p4.metric("Completed / Delivered", delivered_count)

        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader("📊 PO Delivery Pipeline & Aging")
        health_summary = df["Delivery Health"].value_counts().reset_index()
        health_summary.columns = ["Status", "Count"]
        st.bar_chart(health_summary.set_index("Status"), color="#38bdf8", use_container_width=True)

        show_overdue = st.checkbox("🚨 Filter: Show Only Overdue Items", value=False)
        if show_overdue:
            df = df[df["Delivery Health"] == "OVERDUE"]
    else:
        st.info("No 'Due Date' column found in this file to run delivery tracking.")

# Module 5: Spares Inventory
else:
    loc_candidates = ['CS06', 'CS11', 'CCR', 'TRNG', 'CS02', 'CS03', 'CS04', 'CS05', 'CS07', 'CS09', 'CS10']
    avail_locs = [c for c in loc_candidates if c in df.columns]

    for c in avail_locs:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)

    if avail_locs and "TOTAL" not in df.columns:
        df["TOTAL"] = df[avail_locs].sum(axis=1)

    s1, s2, s3 = st.columns(3)
    s1.metric("Registered Spares", len(df))
    if "TOTAL" in df.columns:
        s2.metric("Total Physical Units", f"{int(df['TOTAL'].sum()):,}")
        zero_cnt = int((df["TOTAL"] == 0).sum())
        s3.metric("Zero Stock Deficits", zero_cnt, delta="-Alert" if zero_cnt > 0 else "Normal", delta_color="inverse")

    st.markdown("<br>", unsafe_allow_html=True)
    if avail_locs:
        st.subheader("📊 Station Spares Allocation")
        stn_df = pd.DataFrame({"Station": avail_locs, "Units": [int(df[c].sum()) for c in avail_locs]})
        st.bar_chart(stn_df.set_index("Station"), color="#38bdf8", use_container_width=True)

st.divider()

# Editable Matrix & Write Back
st.subheader(f"Editable Matrix — {active_module}")

search_txt = st.text_input("Search Line Items", placeholder="Search by tag, device, PO, item description...", label_visibility="collapsed")
filtered_df = df.copy()

if search_txt:
    mask = filtered_df.astype(str).apply(lambda r: r.str.contains(search_txt, case=False).any(), axis=1)
    filtered_df = filtered_df[mask]

edited_df = st.data_editor(
    filtered_df,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True
)

st.markdown("<br>", unsafe_allow_html=True)

btn_col1, btn_col2 = st.columns([2, 1])
with btn_col1:
    if st.button(f"💾 Save Modifications Directly to {current_file}", type="primary", use_container_width=True):
        if search_txt:
            df.update(edited_df)
            save_payload = df
        else:
            save_payload = edited_df

        if "Delivery Health" in save_payload.columns:
            save_payload = save_payload.drop(columns=["Delivery Health"])

        save_payload.to_excel(current_file, index=False)
        st.cache_data.clear()
        st.success(f"Workbook '{current_file}' successfully saved to disk!")
        st.rerun()

with btn_col2:
    st.download_button(
        label=f"Export {active_module} (CSV)",
        data=edited_df.to_csv(index=False),
        file_name=f"{selected_config['type']}_export.csv",
        mime="text/csv",
        use_container_width=True
    )
