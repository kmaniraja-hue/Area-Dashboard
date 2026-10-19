import streamlit as st
import pandas as pd
from datetime import datetime, date
import os

# 1. Page Configuration
st.set_page_config(
    page_title="Central Plant C&I Asset & Procurement Portal",
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

# 3. Master Plant Register Mapping
PLANT_MODULES = {
    "📦 DCS Spares & Hardware": {
        "file": "inventory_master.xlsx",
        "desc": "Control room spares allocation, processors, I/O modules, cables.",
        "type": "inventory"
    },
    "⚡ PMIO & Power Supplies (POPS)": {
        "file": "power_supplies_pops.xlsx",
        "desc": "Power supply modules, replacement logs, indigenous vs OEM tracking.",
        "type": "pops"
    },
    "🚰 Valves & COFF Master": {
        "file": "valves_coff.xlsx",
        "desc": "Control valves, on/off cut-off valves, line sizes, area locations.",
        "type": "valves"
    },
    "📝 PR (Indent) Tracker": {
        "file": "pr_tracker.xlsx",
        "desc": "Requisition stages, limited tender justifications, department indents.",
        "type": "pr"
    },
    "🚚 PO Delivery & Due Dates": {
        "file": "po_delivery_status.xlsx",
        "desc": "Vendor tracking, PO delivery due dates, aging, and overdue alarms.",
        "type": "po"
    }
}

# 4. Sidebar Navigation
st.sidebar.markdown("## 🏭 Plant Asset Hub")
st.sidebar.caption("Central Refinery C&I Unit")

active_module = st.sidebar.radio(
    "Select Plant Ledger:",
    list(PLANT_MODULES.keys())
)

selected_config = PLANT_MODULES[active_module]
current_file = selected_config["file"]

st.sidebar.divider()
st.sidebar.markdown(f"**Target File:** `{current_file}`")

# 5. Generic Excel Loader Function
@st.cache_data
def load_excel_sheet(filename):
    if not os.path.exists(filename):
        return None
    try:
        raw = pd.read_excel(filename)
        for idx, row in raw.head(10).iterrows():
            if "S.NO." in row.values or "PO NO" in row.values or "TAG" in row.values or "PR NO" in row.values:
                raw.columns = raw.iloc[idx]
                raw = raw.iloc[idx + 1:].reset_index(drop=True)
                break
        return raw.dropna(how="all")
    except Exception as e:
        st.error(f"Error loading {filename}: {e}")
        return None

df = load_excel_sheet(current_file)

# Starter Template Generator if file is missing
if df is None:
    st.warning(f"File `{current_file}` was not found in the project directory.")
    if st.button(f"Generate starter template for {active_module}"):
        if selected_config["type"] == "po":
            sample_df = pd.DataFrame([
                {"PO NO": "PO-2026-001", "Vendor": "Honeywell", "Item": "C300 Controllers", "PO Date": "2026-08-01", "Due Date": "2026-10-15", "Status": "Pending", "Qty": 4},
                {"PO NO": "PO-2026-002", "Vendor": "Indigenous Power Inc", "Item": "PMIO Power Supplies", "PO Date": "2026-09-10", "Due Date": "2026-11-20", "Status": "Delivered", "Qty": 7},
                {"PO NO": "PO-2026-003", "Vendor": "Flowserve India", "Item": "16-inch 90-deg On/Off Valves", "PO Date": "2026-05-12", "Due Date": "2026-09-30", "Status": "Pending", "Qty": 2}
            ])
        elif selected_config["type"] == "pr":
            sample_df = pd.DataFrame([
                {"PR NO": "PR-C&I-2026-101", "Description": "HCL Tank Radar Level Transmitters Area 36", "Estimated Value (INR)": 850000, "Indent Date": "2026-09-01", "Stage": "Technical Clarification"},
                {"PR NO": "PR-C&I-2026-102", "Description": "12-inch Steam Control Valves Phase-4", "Estimated Value (INR)": 2400000, "Indent Date": "2026-05-20", "Stage": "LTE Sanctioned"}
            ])
        elif selected_config["type"] == "pops":
            sample_df = pd.DataFrame([
                {"Tag No": "PSU-CS09-01", "Station": "CS09", "Make": "Indigenous", "Installed Date": "2026-04-10", "Health": "Good", "Backup Spares": 3},
                {"Tag No": "PSU-CS02-02", "Station": "CS02", "Make": "Honeywell", "Installed Date": "2022-01-15", "Health": "Aging", "Backup Spares": 1}
            ])
        elif selected_config["type"] == "valves":
            sample_df = pd.DataFrame([
                {"Tag No": "XV-36-001", "Area": "Area 36", "Line Size": "16-inch", "Service": "HRD-1/2 Feed", "Type": "COFF On/Off", "Health": "Healthy"},
                {"Tag No": "FV-40-012", "Area": "Phase-4", "Line Size": "12-inch", "Service": "High Press Steam", "Type": "Control Valve", "Health": "Maintenance Due"}
            ])
        else:
            sample_df = pd.DataFrame([{"Item": "Sample Spare", "Stock": 10, "Location": "CS02"}])
        
        sample_df.to_excel(current_file, index=False)
        st.cache_data.clear()
        st.success("Template created! Reloading...")
        st.rerun()
    st.stop()

# 6. Page Header
st.markdown(f'<div class="plant-header">{active_module}</div>', unsafe_allow_html=True)
st.markdown(f'<div class="plant-sub">{selected_config["desc"]}</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# MODULE 1: PO DELIVERY & DUE / OVERDUE ENGINE
# ------------------------------------------------------------------------------
if selected_config["type"] == "po":
    due_col = [c for c in df.columns if "DUE" in str(c).upper()]
    status_col = [c for c in df.columns if "STATUS" in str(c).upper()]

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
        p1.metric("Total POs Tracked", len(df))
        p2.metric("Critical Overdue POs", overdue_count, delta=f"-{overdue_count} Overdue" if overdue_count > 0 else "Clear", delta_color="inverse")
        p3.metric("Due in Next 15 Days", due_soon_count)
        p4.metric("Completed Deliveries", delivered_count)

        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader("📊 Purchase Order Aging & Delivery Pipeline")
        health_summary = df["Delivery Health"].value_counts().reset_index()
        health_summary.columns = ["Delivery Health", "Count"]
        st.bar_chart(health_summary.set_index("Delivery Health"), color="#38bdf8", use_container_width=True)

        show_overdue = st.checkbox("🚨 Show Only Overdue POs", value=False)
        if show_overdue:
            df = df[df["Delivery Health"] == "OVERDUE"]

# ------------------------------------------------------------------------------
# MODULE 2: PR TRACKER
# ------------------------------------------------------------------------------
elif selected_config["type"] == "pr":
    stage_col = [c for c in df.columns if "STAGE" in str(c).upper() or "STATUS" in str(c).upper()]
    val_col = [c for c in df.columns if "VAL" in str(c).upper() or "EST" in str(c).upper()]

    r1, r2, r3 = st.columns(3)
    r1.metric("Active Indents (PRs)", len(df))
    if val_col:
        total_val = pd.to_numeric(df[val_col[0]], errors="coerce").fillna(0).sum()
        r2.metric("Total Requisition Estimate (INR)", f"₹ {total_val:,.0f}")
    if stage_col:
        r3.metric("Procurement Stages", df[stage_col[0]].nunique())

    st.markdown("<br>", unsafe_allow_html=True)
    if stage_col:
        st.subheader("📊 Requisitions by Procurement Stage")
        stage_counts = df[stage_col[0]].value_counts().reset_index()
        stage_counts.columns = ["Stage", "Indent Count"]
        st.bar_chart(stage_counts.set_index("Stage"), color="#0ea5e9", use_container_width=True)

# ------------------------------------------------------------------------------
# MODULE 3: POWER SUPPLIES (POPS)
# ------------------------------------------------------------------------------
elif selected_config["type"] == "pops":
    make_col = [c for c in df.columns if "MAKE" in str(c).upper()]
    health_col = [c for c in df.columns if "HEALTH" in str(c).upper()]
    stn_col = [c for c in df.columns if "STATION" in str(c).upper() or "ROOM" in str(c).upper()]

    k1, k2, k3 = st.columns(3)
    k1.metric("Total Power Supplies Tracked", len(df))
    if make_col:
        k2.metric("Makes Installed", df[make_col[0]].nunique())
    if health_col:
        aging = int((df[health_col[0]].astype(str).str.upper().str.contains("AGING|DEFECT|FAULT")).sum())
        k3.metric("Units Needing Attention", aging, delta="-Action Needed" if aging > 0 else "Normal", delta_color="inverse")

    st.markdown("<br>", unsafe_allow_html=True)
    if stn_col:
        st.subheader("📊 Power Supply Distribution Across Control Stations")
        stn_dist = df[stn_col[0]].value_counts().reset_index()
        stn_dist.columns = ["Station", "PSU Units"]
        st.bar_chart(stn_dist.set_index("Station"), color="#6366f1", use_container_width=True)

# ------------------------------------------------------------------------------
# MODULE 4: VALVES & COFF
# ------------------------------------------------------------------------------
elif selected_config["type"] == "valves":
    area_col = [c for c in df.columns if "AREA" in str(c).upper() or "PHASE" in str(c).upper()]
    size_col = [c for c in df.columns if "SIZE" in str(c).upper()]

    v1, v2, v3 = st.columns(3)
    v1.metric("Total Valve Assets", len(df))
    if area_col:
        v2.metric("Plant Areas Covered", df[area_col[0]].nunique())
    if size_col:
        v3.metric("Bore / Size Variants", df[size_col[0]].nunique())

    st.markdown("<br>", unsafe_allow_html=True)
    if area_col:
        st.subheader("📊 Valve Population by Area")
        area_chart = df[area_col[0]].value_counts().reset_index()
        area_chart.columns = ["Area", "Valve Count"]
        st.bar_chart(area_chart.set_index("Area"), color="#06b6d4", use_container_width=True)

# ------------------------------------------------------------------------------
# MODULE 5: DCS HARDWARE & SPARES
# ------------------------------------------------------------------------------
else:
    location_cols = ['CS06', 'CS11', 'CCR', 'TRNG', 'CS02', 'CS03', 'CS04', 'CS05', 'CS07', 'CS09', 'CS10']
    avail_locs = [c for c in location_cols if c in df.columns]
    for c in avail_locs:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
    if avail_locs:
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

# ------------------------------------------------------------------------------
# DATA EDITOR & WRITE BACK
# ------------------------------------------------------------------------------
st.subheader(f"Editable Data Matrix — {active_module}")

search_txt = st.text_input("Quick Search this Register", placeholder="Search by tag, vendor, PO, or description...")
filtered_df = df.copy()
if search_txt:
    mask = filtered_df.astype(str).apply(lambda row: row.str.contains(search_txt, case=False).any(), axis=1)
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
    if st.button(f"💾 Save All Modifications Directly to {current_file}", type="primary", use_container_width=True):
        if search_txt:
            df.update(edited_df)
            save_payload = df
        else:
            save_payload = edited_df

        if "Delivery Health" in save_payload.columns:
            save_payload = save_payload.drop(columns=["Delivery Health"])

        save_payload.to_excel(current_file, index=False)
        st.cache_data.clear()
        st.success(f"Workbook '{current_file}' successfully updated on disk!")
        st.rerun()

with btn_col2:
    st.download_button(
        label=f"Export {active_module} (CSV)",
        data=edited_df.to_csv(index=False),
        file_name=f"{selected_config['type']}_export.csv",
        mime="text/csv",
        use_container_width=True
    )
