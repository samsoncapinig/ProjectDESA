import streamlit as st
import pandas as pd
import plotly.express as px
import re
import io

st.set_page_config(page_title="Daily Evaluation Summarizer App", layout="wide")
st.title("📊 Project DESA (Daily Evaluation Summarizer App)")
st.caption("Project DESA (Daily Evaluation Summarizer App) is developed by the SMME Section of the SDO Masbate City.")


# -------------------------
# Helpers
# -------------------------
@st.cache_data
def load_file(uploaded_file):
    if uploaded_file.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
    return df.loc[:, ~df.columns.str.contains(r'^Unnamed', regex=True)]

def categorize_columns(df):
    categories = {
        "PROGRAM MANAGEMENT": [],
        "TRAINING VENUE": [],
        "FOOD/MEALS": [],
        "ACCOMMODATION": [],
        "ADMINISTRATIVE ARRANGEMENTS": [],
        "SESSION": []
    }
    for col in df.columns:
        col_str = str(col).upper()
        if "PROGRAM MANAGEMENT" in col_str:
            categories["PROGRAM MANAGEMENT"].append(col)
        elif "TRAINING VENUE" in col_str:
            categories["TRAINING VENUE"].append(col)
        elif "FOOD/MEALS" in col_str:
            categories["FOOD/MEALS"].append(col)
        elif "ACCOMMODATION" in col_str:
            categories["ACCOMMODATION"].append(col)
        elif "ADMINISTRATIVE ARRANGEMENTS" in col_str:
            categories["ADMINISTRATIVE ARRANGEMENTS"].append(col)
        elif any(key in col_str for key in [
            "PROGRAM OBJECTIVES", "LR MATERIALS",
            "CONTENT RELEVANCE", "RP/SUBJECT MATTER EXPERT KNOWLEDGE"
        ]):
            categories["SESSION"].append(col)
    return categories

def compute_category_averages(df, categories, file_name):
    stats = {}
    for cat in ["PROGRAM MANAGEMENT", "TRAINING VENUE", "FOOD/MEALS", "ACCOMMODATION","ADMINISTRATIVE ARANGEMENTS"]:
        cols = categories.get(cat, [])
        if not cols:
            continue
        sub = df[cols].apply(pd.to_numeric, errors='coerce')
        stacked = sub.stack()
        avg = float(stacked.mean()) if not stacked.empty else float("nan")
        stats[cat] = {f"{file_name}": round(avg, 2) if pd.notna(avg) else None}
    return pd.DataFrame(stats).T if stats else None

def compute_session_averages(df, session_cols, file_name):
    session_groups = {}
    for col in session_cols:
        col_str = str(col)
        match = re.search(r"Q\d+[_-]?\s*DAY\s*\d+\s*[-–]?\s*LM\s*\d+", col_str, re.IGNORECASE)
        if match:
            session_key = match.group(0).upper().replace(" ", "")
        else:
            match_day_lm = re.search(r"DAY\s*\d+\s*[-–]?\s*LM\s*\d+", col_str, re.IGNORECASE)
            session_key = match_day_lm.group(0).upper().replace(" ", "") if match_day_lm else None

        if session_key:
            session_groups.setdefault(session_key, []).append(col)
        else:
            st.warning(f"Skipped column (no session match): {col}")

    if not session_groups:
        return None

    stats = {}
    for session, cols in session_groups.items():
        sub = df[cols].apply(pd.to_numeric, errors='coerce')
        stacked = sub.stack()
        avg = float(stacked.mean()) if not stacked.empty else float("nan")
        stats[session] = {f"{file_name}": round(avg, 2) if pd.notna(avg) else None}
    return pd.DataFrame(stats).T if stats else None

def style_numeric_columns(df):
    return df.style.format("{:.2f}")

def add_overall_summary(df):
    """Add a 'Grand Average' column across all files."""
    overall_avg = df.mean(axis=1)
    df["Overall Avg"] = overall_avg.round(2)
    return df

def make_csv_download(df, filename="summary.csv"):
    """Generate CSV file download button for a dataframe."""
    buffer = io.StringIO()
    df.to_csv(buffer)
    st.download_button(
        label=f"⬇️ Download {filename}",
        data=buffer.getvalue(),
        file_name=filename,
        mime="text/csv"
    )

# -------------------------
# Main App
# -------------------------
uploaded_files = st.file_uploader(
    "📂 Upload one or more CSV/XLSX files",
    type=["csv", "xlsx"],
    accept_multiple_files=True
)

if uploaded_files is not None and len(uploaded_files) > 0:
    combined_summary = []
    combined_sessions = []

    for uploaded_file in uploaded_files:
        try:
            df = load_file(uploaded_file)
        except Exception as e:
            st.error(f"Could not read {uploaded_file.name}: {e}")
            continue

        categories = categorize_columns(df)

        cat_df = compute_category_averages(df, categories, uploaded_file.name)
        if cat_df is not None:
            combined_summary.append(cat_df)

        sess_df = compute_session_averages(df, categories.get("SESSION", []), uploaded_file.name)
        if sess_df is not None:
            combined_sessions.append(sess_df)

    # ---- Category comparison ----
    if combined_summary:
        final_summary = pd.concat(combined_summary, axis=1).sort_index()
        final_summary = add_overall_summary(final_summary)
        final_summary.index.name = "Category"

        st.subheader("📌 Category Averages Comparison")
        st.dataframe(style_numeric_columns(final_summary))
        make_csv_download(final_summary, filename="Category_Summary.csv")

        fig = px.bar(
            final_summary.reset_index().melt(
                id_vars="Category",
                var_name="File",
                value_name="Average"
            ),
            x="Category",
            y="Average",
            color="File",
            barmode="group",
            title="Category averages by file"
        )
        st.plotly_chart(fig)

    # ---- Session comparison ----
    if combined_sessions:
        final_sessions = pd.concat(combined_sessions, axis=1).sort_index()
        final_sessions = add_overall_summary(final_sessions)
        final_sessions.index.name = "Session"

        st.subheader("📌 Session Averages Comparison")
        st.dataframe(style_numeric_columns(final_sessions))
        make_csv_download(final_sessions, filename="Session_Summary.csv")

        fig = px.bar(
            final_sessions.reset_index().melt(
                id_vars="Session",
                var_name="File",
                value_name="Average"
            ),
            x="Session",
            y="Average",
            color="File",
            barmode="group",
            title="Session averages by file"
        )
        st.plotly_chart(fig)
else:
    st.info("Upload one or more CSV/XLSX files to generate comparison tables.")
