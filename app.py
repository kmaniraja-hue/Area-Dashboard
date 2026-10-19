import json
import os
import gspread
import pandas as pd
import streamlit as st

# 1. Page Configuration
st.set_page_config(
    page_title="Plant Asset & Operations Portal",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

# 2. Industrial Theme Styling
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


# 3. Google Sheets Authentication
@st.cache_resource
def get_gspread_client():
  service_account_str = os.environ.get("GCP_SERVICE_ACCOUNT")
  if not service_account_str:
    st.error("GCP_SERVICE_ACCOUNT environment variable is not configured.")
    st.stop()
  creds_dict = json.loads(service_account_str)
  return gspread.service_account_from_dict(creds_dict)


gc = get_gspread_client()

# 4. Sheet Registers Configuration
# (These names must match your Google Sheet titles in Google Drive)
REGISTERS = {
    "📝 PR and PO Tracker": "PR and PO",
    "⚡ PMPS Schedule": "PMPS",
    "📦 Spares Inventory": "Spares",
    "🌐 IP Address Register": "IP address",
    "👥 C-Off & Attendance": "Coff Status",
}

st.sidebar.markdown("## 🏭 Plant Operations Hub")
st.sidebar.caption("Google Sheets Cloud Sync Active")

active_label = st.sidebar.radio("Select Register:", list(REGISTERS.keys()))
sheet_title = REGISTERS[active_label]

try:
  sh = gc.open(sheet_title)
except Exception as e:
  st.error(
      f"Unable to access Google Sheet '{sheet_title}'. Ensure it is shared with"
      " your service account email as Editor."
  )
  st.stop()

# Select Worksheet inside the Spreadsheet
worksheets = [ws.title for ws in sh.worksheets()]
selected_worksheet_title = st.sidebar.selectbox("Select Sheet:", worksheets)
ws = sh.worksheet(selected_worksheet_title)


# 5. Smart Universal Data Loader
@st.cache_data(ttl=60)
def load_sheet_data(sheet_name, worksheet_name):
  spreadsheet = gc.open(sheet_name)
  worksheet = spreadsheet.worksheet(worksheet_name)
  raw_values = worksheet.get_all_values()
  if not raw_values:
    return pd.DataFrame()

  raw = pd.DataFrame(raw_values)

  # Check top 10 rows for table header labels
  header_idx = None
  for idx, row in raw.head(10).iterrows():
    row_str = " ".join([str(v).upper() for v in row.values if v])
    if any(
        k in row_str
        for k in [
            "TAG",
            "DESCRIPTION",
            "INDENT",
            "NODE",
            "EMPLOYEE",
            "MATERIAL",
            "S.NO",
            "PO",
            "ITEM",
            "QTY",
        ]
    ):
      header_idx = idx
      break

  if header_idx is not None:
    raw.columns = raw.iloc[header_idx]
    raw = raw.iloc[header_idx + 1 :].reset_index(drop=True)
  else:
    raw.columns = raw.iloc[0]
    raw = raw.iloc[1:].reset_index(drop=True)

  raw = raw.dropna(how="all", axis=0).dropna(how="all", axis=1)

  # Clean and deduplicate column names
  clean_cols = []
  seen = {}
  for i, c in enumerate(raw.columns):
    name = (
        f"Col_{i+1}"
        if (pd.isna(c) or str(c).strip() == "" or "UNNAMED:" in str(c).upper())
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


df = load_sheet_data(sheet_title, selected_worksheet_title)

st.markdown(
    f'<div class="plant-header">{active_label} —'
    f" {selected_worksheet_title}</div>",
    unsafe_allow_html=True,
)
st.markdown(
    f'<div class="plant-sub">Source: Google Sheets ({sheet_title}) | Total'
    f" Records: {len(df)}</div>",
    unsafe_allow_html=True,
)

# ------------------------------------------------------------------------------
# 6. KPI METRICS & CHARTS
# ------------------------------------------------------------------------------
numeric_cols = []
for c in df.columns:
  converted = pd.to_numeric(
      df[c].astype(str).str.replace(",", "").str.replace("₹", "").str.strip(),
      errors="coerce",
  )
  if converted.notna().sum() > (len(df) * 0.4) and len(df) > 0:
    numeric_cols.append(c)

text_cols = [c for c in df.columns if c not in numeric_cols]

m1, m2, m3 = st.columns(3)
m1.metric("Total Line Items", len(df))

if numeric_cols:
  primary_num = numeric_cols[0]
  clean_s = pd.to_numeric(
      df[primary_num]
      .astype(str)
      .str.replace(",", "")
      .str.replace("₹", "")
      .str.strip(),
      errors="coerce",
  ).fillna(0)
  m2.metric(f"Total {primary_num}", f"{clean_s.sum():,.0f}")
else:
  m2.metric("Total Data Columns", len(df.columns))

if text_cols:
  primary_cat = text_cols[0]
  m3.metric(f"Unique {primary_cat}", df[primary_cat].dropna().nunique())
else:
  m3.metric("Total Worksheets", len(worksheets))

st.markdown("#### 📊 Visual Summary")
chart_col1, chart_col2 = st.columns(2)

with chart_col1:
  chartable_cat = [
      c
      for c in text_cols
      if df[c].dropna().nunique() > 1 and df[c].dropna().nunique() <= 30
  ]
  if chartable_cat:
    cat = chartable_cat[0]
    st.caption(f"**Distribution by {cat}**")
    cat_data = df[cat].value_counts().head(10).reset_index()
    cat_data.columns = [cat, "Count"]
    st.bar_chart(cat_data.set_index(cat), color="#38bdf8")
  else:
    st.info("No categorical distribution available.")

with chart_col2:
  if numeric_cols and text_cols:
    num = numeric_cols[0]
    lbl = text_cols[0]
    st.caption(f"**Top Items by {num}**")
    temp_df = df[[lbl, num]].copy()
    temp_df[num] = pd.to_numeric(
        temp_df[num]
        .astype(str)
        .str.replace(",", "")
        .str.replace("₹", "")
        .str.strip(),
        errors="coerce",
    ).fillna(0)
    top_data = temp_df.sort_values(by=num, ascending=False).head(10)
    st.bar_chart(top_data.set_index(lbl), color="#10b981")
  else:
    st.info("Numeric breakdown will appear when numeric fields are populated.")

st.divider()

# ------------------------------------------------------------------------------
# 7. SEARCH, EDITABLE GRID & TWO-WAY GOOGLE SHEETS WRITE-BACK
# ------------------------------------------------------------------------------
st.subheader(f"Editable Register — {selected_worksheet_title}")

search_txt = st.text_input(
    "Search Data",
    placeholder="Search across all fields...",
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
      f"☁️ Sync Modifications to Google Sheets ({selected_worksheet_title})",
      type="primary",
      use_container_width=True,
  ):
    try:
      # Convert DataFrame back to 2D list
      export_data = [edited_df.columns.values.tolist()] + edited_df.fillna(
          ""
      ).values.tolist()
      ws.clear()
      ws.update(range_name="A1", values=export_data)
      st.cache_data.clear()
      st.success("Successfully synchronized live to Google Drive!")
      st.rerun()
    except Exception as e:
      st.error(f"Sync failed: {e}")

with btn2:
  st.download_button(
      label=f"Export {selected_worksheet_title} (CSV)",
      data=edited_df.to_csv(index=False),
      file_name=f"{sheet_title}_{selected_worksheet_title}.csv",
      mime="text/csv",
      use_container_width=True,
  )
