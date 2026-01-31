import io
import pandas as pd
import streamlit as st
# pip install streamlit pandas openpyxl xlsxwriter
st.set_page_config(page_title="Attendance Summary", layout="wide")

st.title("📊 Attendance Analyzer (Excel → Present/Absent Summary)")
st.write("Upload an Excel file, map columns, and download a clean summary. Designed for agent-friendly usage ✅")

# -----------------------------
# Helpers
# -----------------------------
DEFAULT_STATUS_MAP = {
    "p": "present",
    "present": "present",
    "pr": "present",
    "1": "present",
    "yes": "present",
    "y": "present",

    "a": "absent",
    "absent": "absent",
    "ab": "absent",
    "0": "absent",
    "no": "absent",
    "n": "absent",

    "leave": "absent",
    "l": "absent",
    "lop": "absent",
}

DEFAULT_IGNORE = {"", "nan", "none", "na", "n/a", "-", "holiday", "week off", "wo", "w/o"}

def normalize_status(val: str, status_map: dict, ignore_set: set):
    if pd.isna(val):
        return None
    s = str(val).strip().lower()
    if s in ignore_set:
        return None
    # remove extra spaces
    s = " ".join(s.split())
    return status_map.get(s, s)  # keep unknown labels as-is

def detect_format(df: pd.DataFrame):
    """
    Heuristic: if there is a 'status' column -> long format.
    Else, assume wide format.
    """
    cols = [c.strip().lower() for c in df.columns]
    if "status" in cols:
        return "Long (Agent/Date/Status)"
    return "Wide (Agent + Date columns)"

def to_excel_bytes(raw_df, long_df, agent_summary_df, overall_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        raw_df.to_excel(writer, index=False, sheet_name="Raw_Data")
        if long_df is not None:
            long_df.to_excel(writer, index=False, sheet_name="Normalized_Long")
        agent_summary_df.to_excel(writer, index=False, sheet_name="Agent_Summary")
        overall_df.to_excel(writer, index=False, sheet_name="Overall_Summary")

        # Basic formatting
        workbook  = writer.book
        header_fmt = workbook.add_format({"bold": True, "bg_color": "#DCE6F1", "border": 1})
        for sheet in ["Raw_Data", "Agent_Summary", "Overall_Summary", "Normalized_Long"]:
            if sheet in writer.sheets:
                ws = writer.sheets[sheet]
                # set columns width a bit
                ws.set_column(0, 50, 18)
                # header formatting
                try:
                    for col_num, value in enumerate(
                        (raw_df.columns if sheet=="Raw_Data"
                         else agent_summary_df.columns if sheet=="Agent_Summary"
                         else overall_df.columns if sheet=="Overall_Summary"
                         else long_df.columns)
                    ):
                        ws.write(0, col_num, value, header_fmt)
                except Exception:
                    pass

    output.seek(0)
    return output.getvalue()

# -----------------------------
# UI: Upload
# -----------------------------
uploaded = st.file_uploader("📥 Upload Attendance Excel (.xlsx/.xls)", type=["xlsx", "xls"])

if not uploaded:
    st.info("Upload an Excel file to begin.")
    st.stop()

# Load workbook
try:
    xl = pd.ExcelFile(uploaded)
    sheet_names = xl.sheet_names
except Exception as e:
    st.error(f"Could not read Excel file. Error: {e}")
    st.stop()

col1, col2 = st.columns([2, 3])
with col1:
    sheet = st.selectbox("Select sheet", sheet_names)

# Read selected sheet
try:
    df = pd.read_excel(uploaded, sheet_name=sheet, engine="openpyxl")
except Exception:
    # fallback (some .xls)
    df = pd.read_excel(uploaded, sheet_name=sheet)

if df.empty:
    st.warning("Selected sheet is empty.")
    st.stop()

# Show preview
st.subheader("🔎 Data Preview")
st.dataframe(df.head(20), use_container_width=True)

# -----------------------------
# Detect or choose format
# -----------------------------
suggested = detect_format(df)
with col2:
    fmt = st.radio(
        "Data layout",
        options=["Long (Agent/Date/Status)", "Wide (Agent + Date columns)"],
        index=0 if suggested.startswith("Long") else 1,
        help="Long: one row per agent per day. Wide: one row per agent, many date columns with P/A."
    )

# -----------------------------
# Status mapping controls
# -----------------------------
st.subheader("⚙️ Status Normalization Rules")

left, right = st.columns([2, 2])
with left:
    st.caption("You can customize mapping if your sheet uses different codes.")
    custom_map_text = st.text_area(
        "Status map (one per line: raw=value)",
        value="\n".join([f"{k}={v}" for k, v in DEFAULT_STATUS_MAP.items()]),
        height=200
    )
with right:
    ignore_text = st.text_area(
        "Ignore values (one per line)",
        value="\n".join(sorted(DEFAULT_IGNORE)),
        height=200
    )

# Parse custom map
status_map = {}
for line in custom_map_text.splitlines():
    line = line.strip()
    if not line or "=" not in line:
        continue
    k, v = line.split("=", 1)
    status_map[k.strip().lower()] = v.strip().lower()

ignore_set = set([x.strip().lower() for x in ignore_text.splitlines() if x.strip()])

# -----------------------------
# Processing based on format
# -----------------------------
st.subheader("✅ Processing")

if fmt == "Long (Agent/Date/Status)":
    # Column selection
    cols = list(df.columns)
    lower_map = {c.strip().lower(): c for c in cols}

    agent_default = lower_map.get("agent") or lower_map.get("agent name") or lower_map.get("name") or cols[0]
    status_default = lower_map.get("status") or lower_map.get("attendance") or cols[-1]
    date_default = lower_map.get("date") or lower_map.get("day") or None

    c1, c2, c3 = st.columns(3)
    with c1:
        agent_col = st.selectbox("Agent column", cols, index=cols.index(agent_default) if agent_default in cols else 0)
    with c2:
        status_col = st.selectbox("Status column", cols, index=cols.index(status_default) if status_default in cols else len(cols)-1)
    with c3:
        date_col = st.selectbox("Date column (optional)", ["(none)"] + cols, index=0 if date_default is None else (cols.index(date_default)+1))

    # Normalize
    work = df.copy()
    work[agent_col] = work[agent_col].astype(str).str.strip()

    norm_status = work[status_col].apply(lambda x: normalize_status(x, status_map, ignore_set))
    work["_status_norm"] = norm_status

    # Filter only present/absent (and any other statuses will stay, you can decide)
    # We'll count only present/absent by default:
    filtered = work.dropna(subset=["_status_norm"]).copy()

    # Overall counts
    overall_counts = filtered["_status_norm"].value_counts()
    total_present = int(overall_counts.get("present", 0))
    total_absent = int(overall_counts.get("absent", 0))
    total_records = int(len(filtered))

    overall_df = pd.DataFrame({
        "Metric": ["Total Present", "Total Absent", "Total Counted Records"],
        "Value": [total_present, total_absent, total_records]
    })

    # Per agent
    agent_summary = (
        filtered.pivot_table(index=agent_col, columns="_status_norm", aggfunc="size", fill_value=0)
        .reset_index()
    )
    if "present" not in agent_summary.columns:
        agent_summary["present"] = 0
    if "absent" not in agent_summary.columns:
        agent_summary["absent"] = 0

    agent_summary["total_counted"] = agent_summary["present"] + agent_summary["absent"]
    agent_summary["attendance_%"] = agent_summary.apply(
        lambda r: round((r["present"] / r["total_counted"] * 100), 2) if r["total_counted"] else 0.0,
        axis=1
    )
    agent_summary = agent_summary.sort_values(["attendance_%", "total_counted"], ascending=[False, False])

    st.success("Done! See summaries below 👇")

    a1, a2 = st.columns(2)
    with a1:
        st.subheader("Overall Summary")
        st.dataframe(overall_df, use_container_width=True)
    with a2:
        st.subheader("Per-Agent Summary")
        st.dataframe(agent_summary, use_container_width=True)

    # Optional chart
    st.subheader("📈 Quick Chart")
    chart_df = agent_summary[[agent_col, "present", "absent"]].set_index(agent_col)
    st.bar_chart(chart_df)

    # Download
    excel_bytes = to_excel_bytes(df, filtered.drop(columns=[]), agent_summary, overall_df)
    st.download_button(
        label="⬇️ Download Summary Excel",
        data=excel_bytes,
        file_name="attendance_summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    # Wide format
    cols = list(df.columns)
    agent_col = st.selectbox("Agent column (usually first column)", cols, index=0)

    # Choose date columns
    date_cols = [c for c in cols if c != agent_col]
    selected_date_cols = st.multiselect(
        "Select attendance/date columns (P/A values)",
        options=date_cols,
        default=date_cols[:],
        help="These columns contain P/A or Present/Absent values."
    )

    if not selected_date_cols:
        st.warning("Select at least one attendance/date column.")
        st.stop()

    work = df.copy()
    work[agent_col] = work[agent_col].astype(str).str.strip()

    # Melt to long
    long_df = work.melt(id_vars=[agent_col], value_vars=selected_date_cols,
                        var_name="date", value_name="status_raw")

    long_df["_status_norm"] = long_df["status_raw"].apply(lambda x: normalize_status(x, status_map, ignore_set))
    long_df = long_df.dropna(subset=["_status_norm"])

    overall_counts = long_df["_status_norm"].value_counts()
    total_present = int(overall_counts.get("present", 0))
    total_absent = int(overall_counts.get("absent", 0))
    total_records = int(len(long_df))

    overall_df = pd.DataFrame({
        "Metric": ["Total Present", "Total Absent", "Total Counted Records"],
        "Value": [total_present, total_absent, total_records]
    })

    agent_summary = (
        long_df.pivot_table(index=agent_col, columns="_status_norm", aggfunc="size", fill_value=0)
        .reset_index()
    )
    if "present" not in agent_summary.columns:
        agent_summary["present"] = 0
    if "absent" not in agent_summary.columns:
        agent_summary["absent"] = 0

    agent_summary["total_counted"] = agent_summary["present"] + agent_summary["absent"]
    agent_summary["attendance_%"] = agent_summary.apply(
        lambda r: round((r["present"] / r["total_counted"] * 100), 2) if r["total_counted"] else 0.0,
        axis=1
    )
    agent_summary = agent_summary.sort_values(["attendance_%", "total_counted"], ascending=[False, False])

    st.success("Done! See summaries below 👇")

    a1, a2 = st.columns(2)
    with a1:
        st.subheader("Overall Summary")
        st.dataframe(overall_df, use_container_width=True)
    with a2:
        st.subheader("Per-Agent Summary")
        st.dataframe(agent_summary, use_container_width=True)

    st.subheader("📈 Quick Chart")
    chart_df = agent_summary[[agent_col, "present", "absent"]].set_index(agent_col)
    st.bar_chart(chart_df)

    # Download
    excel_bytes = to_excel_bytes(df, long_df, agent_summary, overall_df)
    st.download_button(
        label="⬇️ Download Summary Excel",
        data=excel_bytes,
        file_name="attendance_summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.caption("Tip: Add more mappings like WFH=present, Half Day=present etc. in the Status map box.")