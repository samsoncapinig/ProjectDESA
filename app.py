import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import tempfile
from typing import Dict, List, Tuple, Optional

# Plotting
import plotly.express as px

# PPTX
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# PDF
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch

# =============================
# PAGE CONFIG
# =============================
st.set_page_config(page_title="Evaluation Dashboard", layout="wide")

st.title("📊 Evaluation Dashboard (Combined Version)")
st.caption("Auto-detects CSV / Excel files | Generates summaries, insights, and reports")

# =============================
# UNIVERSAL FILE LOADER
# =============================
def load_any_file(uploaded_file):
    try:
        return pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception:
        try:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file)
        except Exception as e:
            st.error(f"❌ Unsupported file: {e}")
            return None

# =============================
# HELPERS
# =============================
def coerce_numeric(df):
    return df.apply(pd.to_numeric, errors="coerce")

def compute_avg(df, cols):
    if not cols:
        return np.nan
    return coerce_numeric(df[cols]).stack().mean()

# =============================
# INSIGHTS DETECTION
# =============================
INSIGHT_REGEX = re.compile(r"insight", re.IGNORECASE)

def find_insight_cols(df):
    return [c for c in df.columns if INSIGHT_REGEX.search(str(c))]

def extract_insights(df):
    cols = find_insight_cols(df)
    positives, improvements = [], []

    for col in cols:
        for val in df[col].dropna():
            text = str(val).lower()

            if any(w in text for w in ["good", "excellent", "great", "helpful"]):
                positives.append(val)
            elif any(w in text for w in ["improve", "should", "need", "lack"]):
                improvements.append(val)

    return positives, improvements

# =============================
# FILE UPLOAD
# =============================
uploaded_files = st.file_uploader(
    "Upload CSV or Excel Files",
    type=["csv", "xlsx", "xls"],
    accept_multiple_files=True
)

# =============================
# MAIN PROCESS
# =============================
if uploaded_files:

    all_category = []
    all_session = []

    insights_all = {"positive": [], "improve": []}

    for file in uploaded_files:

        st.divider()
        st.subheader(f"📄 {file.name}")

        df = load_any_file(file)

        if df is None:
            continue

        st.success("File loaded successfully")

        # =============================
        # NUMERIC DETECTION
        # =============================
        numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()

        rating_cols = [
            col for col in numeric_cols
            if not any(x in col.lower() for x in ["id", "response"])
        ]

        if rating_cols:
            overall_avg = df[rating_cols].mean().mean()
            st.metric("Overall Rating", round(overall_avg, 2))

        # =============================
        # CATEGORY SUMMARY
        # =============================
        category_map = {}
        for col in rating_cols:
            category = col.split("->")[0] if "->" in col else col
            category_map[col] = category

        category_df = pd.DataFrame({
            "Category": [category_map[c] for c in rating_cols],
            "Average": [df[c].mean() for c in rating_cols]
        })

        category_avg = category_df.groupby("Category").mean().reset_index()
        category_avg["File"] = file.name

        st.dataframe(category_avg)
        st.bar_chart(category_avg.set_index("Category"))

        all_category.append(category_avg)

        # =============================
        # QUALITATIVE FEEDBACK
        # =============================
        st.subheader("📝 Qualitative Feedback")

        qual_cols = [
            "Q12_Most Significant Learning",
            "Q13_Learnings",
            "Q14_Suggestions"
        ]

        qual_cols = [c for c in qual_cols if c in df.columns]

        if qual_cols:
            st.dataframe(df[qual_cols].dropna(how="all"))

        # =============================
        # INSIGHTS
        # =============================
        pos, imp = extract_insights(df)

        insights_all["positive"] += pos
        insights_all["improve"] += imp

    # =============================
    # COMBINED INSIGHTS
    # =============================
    st.divider()
    st.subheader("🗣️ Combined Insights")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**✅ Positive**")
        st.write(insights_all["positive"] or "None")

    with col2:
        st.markdown("**🛠️ For Improvement**")
        st.write(insights_all["improve"] or "None")

    # =============================
    # PDF GENERATOR
    # =============================
    st.subheader("📄 Generate PDF Report")

    if st.button("Generate PDF"):

        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        styles = getSampleStyleSheet()

        elements = []

        elements.append(Paragraph("Evaluation Report", styles["Title"]))
        elements.append(Spacer(1, 12))

        elements.append(Paragraph("Insights", styles["Heading2"]))

        for s in insights_all["positive"]:
            elements.append(Paragraph(f"✔ {s}", styles["Normal"]))

        for s in insights_all["improve"]:
            elements.append(Paragraph(f"⚠ {s}", styles["Normal"]))

        doc.build(elements)
        buffer.seek(0)

        st.download_button("Download PDF", buffer, "report.pdf")

    # =============================
    # PPT GENERATOR
    # =============================
    st.subheader("📊 Generate PPT")

    if st.button("Generate PPT"):

        prs = Presentation()

        slide = prs.slides.add_slide(prs.slide_layouts[5])
        title = slide.shapes.title
        title.text = "Evaluation Report"

        textbox = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(6), Inches(4))
        tf = textbox.text_frame

        tf.text = "Insights"

        for s in insights_all["positive"]:
            p = tf.add_paragraph()
            p.text = f"✔ {s}"

        for s in insights_all["improve"]:
            p = tf.add_paragraph()
            p.text = f"⚠ {s}"

        ppt_buffer = BytesIO()
        prs.save(ppt_buffer)
        ppt_buffer.seek(0)

        st.download_button(
            "Download PPT",
            ppt_buffer,
            "report.pptx"
        )

else:
    st.info("Upload files to start.")
