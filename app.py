import streamlit as st
import pandas as pd
from datetime import datetime, date
import os

# 1. Page Configuration
st.set_page_config(
    page_title="Area-1 C&I Plant Hub",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. Industrial Dark Theme
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
    unsafe_allow_html=True
)

# 3. Module Configurations
PLANT_MODULES = {
    "📝 PR and PO Tracker": {
        "file": "PR and PO.xlsx",
        "desc": "Procurement indents, LTE/Open status, PO awards, and delivery due aging.",
        "type": "pr_po"
    },
    "⚡ PMPS Schedule": {
        "file": "PMPS.xlsx",
        "desc": "Preventive maintenance, calibration frequency, and area instrument distribution.",
        "type": "pmps"
    },
    "📦 Spares Inventory": {
        "file": "Spares.xlsx",
        "desc": "Installed equipment quantities, stock in store, and zero-inventory alerts.",
        "type": "spares"
    },
    "🌐 IP Address Register": {
        "file": "IP address.xlsx",
        "desc": "DCS server nodes, operator stations, subnets, and location allocations.",
        "type": "ip"
    },
    "👥 C-Off & Attendance": {
        "file": "Coff Status.xlsx",
        "desc": "Compensatory off balances, extra duty days, and monthly attendance tracking.",
        "type": "coff"
    }
}

st.sidebar.markdown("## 🏭 Plant Operations Hub")
st.sidebar.caption("Area-1 Instrumentation & Control")

active_module = st.sidebar.radio("Select Register:", list(PLANT_MODULES.keys()))
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
st.sidebar.markdown(f"**Active File:** `{current_file}`")

# 4. Universal Smart Loader
@st.cache_data
def load_clean_data(filename, sheet):
    raw = pd.read_excel(filename, sheet_name=sheet)
    
    # Check top 10 rows for header keywords
    header_idx = None
    for idx, row in raw.head(10).iterrows():
        row_str = " ".join([str(v).upper() for v in row.values if pd.notna(v)])
        if any(k in row_str for k in ["TAG NO", "DESCRIPTION", "INDENT", "NODE NAME", "EMPLOYEE", "MATERIAL CODE", "S.NO"]):
            header_idx = idx
            break

    if header_idx is not None:
        raw.columns = raw.iloc[header_idx]
        raw = raw.iloc[header_idx + 1:].reset_index(drop=True)

    raw = raw.dropna(how="all", axis=0).dropna(how="all", axis=1)

    clean_cols = []
    seen = {}
    for i, c in enumerate(raw.columns):
        name = f"Col_{i+1}" if (pd.isna(c) or str(c).strip() == "" or "UNNAMED:" in str(c).upper()) else str(c).strip()
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        clean_cols.append(name)
    raw.columns = clean_cols
    return raw.reset_index(drop=True)

df = load_clean_data(current_file, selected_sheet)

st.markdown(f'<div class="plant-header">{active_module} — {selected_sheet}</div>', unsafe_allow_html=True)
st.markdown(f'<div class="plant-sub">{selected_config["desc"]}</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# GRAPHICAL REPRESENTATIONS & METRICS
# ------------------------------------------------------------------------------

# --- MODULE 1: PR & PO TRACKER ---
if selected_config["type"] == "pr_po":
    type_col = [c for c in df.columns if "TYPE" in str(c).upper()]
    indentor_col = [c for c in df.columns if "INDENTOR" in str(c).upper()]
    est_col = [c for c in df.columns if "EST" in str(c).upper() or "VALUE" in str(c).upper()]
    po_col = [c for c in df.columns if "PO NO" in str(c).upper()]

    m1, m2, m3 = st.columns(3)
    m1.metric("Total Items Listed", len(df))
    if po_col:
        m2.metric("POs Generated", int(df[po_col[0]].notna().sum()))
    if est_col:
        clean_val = df[est_col[0]].astype(str).str.replace("₹", "").str.replace(",", "").str.strip()
        tot_val = pd.to_numeric(clean_val, errors="coerce").sum()
        m3.metric("Total Est. Requisition", f"₹ {tot_val:,.0f}")

    st.markdown("#### 📊 Procurement Analytics")
    g1, g2 = st.columns(2)
    with g1:
        if type_col and df[type_col[0]].notna().any():
            st.caption("**Procurement Type Breakdown (LTE / Open / Prop)**")
            t_counts = df[type_col[0]].dropna().value_counts().reset_index()
            t_counts.columns = ["Procurement Type", "Count"]
            st.bar_chart(t_counts.set_index("Procurement Type"), color="#38bdf8")
    with g2:
        if indentor_col and df[indentor_col[0]].notna().any():
            st.caption("**Indents by Initiator / Cell**")
            i_counts = df[indentor_col[0]].dropna().value_counts().reset_index()
            i_counts.columns = ["Indentor", "Count"]
            st.bar_chart(i_counts.set_index("Indentor"), color="#6366f1")

# --- MODULE 2: PMPS SCHEDULE ---
elif selected_config["type"] == "pmps":
    area_col = [c for c in df.columns if "AREA" in str(c).upper()]
    crit_col = [c for c in df.columns if "CRIT" in str(c).upper()]
    act_col = [c for c in df.columns if "ACT" in str(c).upper()]

    m1, m2, m3 = st.columns(3)
    m1.metric("Instrumentation Tags", len(df))
    if crit_col:
        crit_e = int((df[crit_col[0]].astype(str).str.strip() == "E").sum())
        m2.metric("Essential (Critical-E) Tags", crit_e, delta="-High Priority" if crit_e > 0 else "Normal", delta_color="inverse")
    if area_col:
        m3.metric("Refinery Process Areas", df[area_col[0]].dropna().nunique())

    st.markdown("#### 📊 PMPS Instrumentation Distribution")
    g1, g2 = st.columns(2)
    with g1:
        if area_col and df[area_col[0]].notna().any():
            st.caption("**Instrumentation Density by Refinery Area**")
            a_counts = df[area_col[0]].dropna().value_counts().reset_index()
            a_counts.columns = ["Area", "Tags"]
            st.bar_chart(a_counts.set_index("Area"), color="#06b6d4")
    with g2:
        if crit_col and df[crit_col[0]].notna().any():
            st.caption("**Tag Criticality Breakdown (E vs Non-E)**")
            c_counts = df[crit_col[0]].dropna().value_counts().reset_index()
            c_counts.columns = ["Criticality", "Count"]
            st.bar_chart(c_counts.set_index("Criticality"), color="#f59e0b")

# --- MODULE 3: SPARES INVENTORY ---
elif selected_config["type"] == "spares":
    stock_col = [c for c in df.columns if "STOCK" in str(c).upper()]
    inst_col = [c for c in df.columns if "INSTL" in str(c).upper()]
    desc_col = [c for c in df.columns if "DESC" in str(c).upper()]

    m1, m2, m3 = st.columns(3)
    m1.metric("Spares Line Items", len(df))

    if stock_col:
        stock_s = pd.to_numeric(df[stock_col[0]], errors="coerce").fillna(0)
        m2.metric("Total Central Store Stock", f"{int(stock_s.sum()):,}")
        zero_cnt = int((stock_s == 0).sum())
        m3.metric("Zero Stock Deficits", zero_cnt, delta="-Attention" if zero_cnt > 0 else "Normal", delta_color="inverse")

    st.markdown("#### 📊 Installed Quantities vs Stock in Store")
    if inst_col and stock_col and desc_col:
        chart_df = df[[desc_col[0], inst_col[0], stock_col[0]]].dropna(subset=[desc_col[0]]).copy()
        chart_df[inst_col[0]] = pd.to_numeric(chart_df[inst_col[0]], errors="coerce").fillna(0)
        chart_df[stock_col[0]] = pd.to_numeric(chart_df[stock_col[0]], errors="coerce").fillna(0)
        # Filter top 15 highest installed assets
        top_chart = chart_df.sort_values(by=inst_col[0], ascending=False).head(15)
        top_chart = top_chart.set_index(desc_col[0])[[inst_col[0], stock_col[0]]]
        top_chart.columns = ["Installed Quantity", "Stock in Store"]
        st.bar_chart(top_chart, color=["#38bdf8", "#10b981"])

# --- MODULE 4: IP ADDRESS REGISTER ---
elif selected_config["type"] == "ip":
    loc_col = [c for c in df.columns if "LOCATION" in str(c).upper() or "LOC" in str(c).upper()]
    mask_col = [c for c in df.columns if "MASK" in str(c).upper() or "SUBNET" in str(c).upper()]

    m1, m2 = st.columns(2)
    m1.metric("Assigned Network Nodes", len(df))
    if loc_col:
        m2.metric("Distinct Control Locations", df[loc_col[0]].dropna().nunique())

    st.markdown("#### 📊 Network Architecture Distribution")
    g1, g2 = st.columns(2)
    with g1:
        if loc_col and df[loc_col[0]].notna().any():
            st.caption("**Nodes Allocated by Location / Control Room**")
            l_counts = df[loc_col[0]].dropna().value_counts().reset_index()
            l_counts.columns = ["Location", "Node Count"]
            st.bar_chart(l_counts.set_index("Location"), color="#38bdf8")
    with g2:
        if mask_col and df[mask_col[0]].notna().any():
            st.caption("**Subnet Mask Segment Allocation**")
            m_counts = df[mask_col[0]].dropna().value_counts().reset_index()
            m_counts.columns = ["Subnet Mask", "Count"]
            st.bar_chart(m_counts.set_index("Subnet Mask"), color="#818cf8")

# --- MODULE 5: COFF & ATTENDANCE ---
else:
    emp_col = [c for c in df.columns if "EMPLOYEE" in str(c).upper() or "NAME" in str(c).upper()]
    co_col = [c for c in df.columns if "CO LEAVE" in str(c).upper() or "CF" in str(c).upper()]
    extra_col = [c for c in df.columns if "EXTRA" in str(c).upper()]

    m1, m2 = st.columns(2)
    m1.metric("Personnel Records", len(df))
    if extra_col:
        tot_extra = int(pd.to_numeric(df[extra_col[0]], errors="coerce").fillna(0).sum())
        m2.metric("Total Extra Working Days", tot_extra)

    st.markdown("#### 📊 C-Off & Extra Days Comparison by Personnel")
    if emp_col and (co_col or extra_col):
        cols_to_use = [emp_col[0]]
        col_names = []
        if co_col:
            cols_to_use.append(co_col[0])
            col_names.append("CO Leave Count")
        if extra_col:
            cols_to_use.append(extra_col[0])
            col_names.append("Extra Working Days")

        c_df = df[cols_to_use].dropna(subset=[emp_col[0]]).copy()
        for c in cols_to_use[1:]:
            c_df[c] = pd.to_numeric(c_df[c], errors="coerce").fillna(0)

        c_df = c_df.set_index(emp_col[0])
        c_df.columns = col_names
        st.bar_chart(c_df, color=["#f43f5e", "#0ea5e9"] if len(col_names) > 1 else ["#0ea5e9"])

st.divider()

# ------------------------------------------------------------------------------
# DATA MATRIX & DIRECT SAVE BACK TO EXCEL
# ------------------------------------------------------------------------------
st.subheader(f"Editable Matrix — {selected_sheet}")

search_txt = st.text_input("Quick Filter", placeholder="Search tags, descriptions, codes...", label_visibility="collapsed")
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

btn1, btn2 = st.columns([2, 1])
with btn1:
    if st.button(f"💾 Save Changes Directly to '{selected_sheet}' in {current_file}", type="primary", use_container_width=True):
        with pd.ExcelWriter(current_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            edited_df.to_excel(writer, sheet_name=selected_sheet, index=False)
        st.cache_data.clear()
        st.success(f"Sheet '{selected_sheet}' successfully saved to disk!")
        st.rerun()

with btn2:
    st.download_button(
        label=f"Export {selected_sheet} (CSV)",
        data=edited_df.to_csv(index=False),
        file_name=f"{selected_sheet}.csv",
        mime="text/csv",
        use_container_width=True
    )
