import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import io

st.set_page_config(page_title="BOM Comparator", layout="wide")

st.title("📊 BOM Comparison Tool")

# ======================
# UPLOAD FILES
# ======================
old_file = st.file_uploader("📂 Upload OLD BOM", type=["xlsx"])
new_file = st.file_uploader("📂 Upload NEW BOM", type=["xlsx"])

# ======================
# BUTTON
# ======================
start = st.button("🚀 Start Comparison")

if start:

    if old_file is None or new_file is None:
        st.error("❌ Please upload both BOM files")
        st.stop()

    # ======================
    # READ FILES
    # ======================
    old_bom = pd.read_excel(old_file)
    new_bom = pd.read_excel(new_file)

    # Clean columns
    old_bom.columns = old_bom.columns.str.strip()
    new_bom.columns = new_bom.columns.str.strip()

    required_cols = ["PN", "Description", "bom_qty", "BOM text"]

    # Check columns
    for col in required_cols:
        if col not in old_bom.columns or col not in new_bom.columns:
            st.error(f"❌ Missing column: {col}")
            st.stop()

    # Keep only needed columns
    old_bom = old_bom[required_cols].copy()
    new_bom = new_bom[required_cols].copy()

    # Rename Position
    old_bom.rename(columns={"BOM text": "Position"}, inplace=True)
    new_bom.rename(columns={"BOM text": "Position"}, inplace=True)

    # ======================
    # CLEAN DATA
    # ======================
    for df in [old_bom, new_bom]:
        df["PN"] = df["PN"].astype(str).str.strip().str.upper()
        df["Position"] = df["Position"].astype(str).str.strip().str.upper()

    old_bom["bom_qty"] = pd.to_numeric(old_bom["bom_qty"], errors="coerce")
    new_bom["bom_qty"] = pd.to_numeric(new_bom["bom_qty"], errors="coerce")

    # ======================
    # CREATE KEY
    # ======================
    old_bom["key"] = old_bom["PN"] + "_" + old_bom["Position"]
    new_bom["key"] = new_bom["PN"] + "_" + new_bom["Position"]

    # ======================
    # MERGE
    # ======================
    df = old_bom.merge(
        new_bom,
        on="key",
        how="outer",
        suffixes=("_old", "_new"),
        indicator=True
    )

    # ======================
    # FILL ONLY TEXT COLUMNS (FIX ERROR)
    # ======================
    text_cols = df.select_dtypes(include=["object"]).columns
    df[text_cols] = df[text_cols].fillna("")

    # ======================
    # STATUS FUNCTION
    # ======================
    def get_status(row):

        if row["_merge"] == "both":
            if row["bom_qty_old"] == row["bom_qty_new"]:
                return "Conforme"
            else:
                return "Qty Difference"

        elif row["_merge"] == "left_only":
            return "Missing in BOM 2"

        elif row["_merge"] == "right_only":
            return "Missing in BOM 1"

        return "Unknown"

    df["Status"] = df.apply(get_status, axis=1)

    # ======================
    # FINAL RESULT TABLE
    # ======================
    result = df[[
        "PN_old", "Description_old", "bom_qty_old",
        "PN_new", "Description_new", "bom_qty_new",
        "Status"
    ]]

    # Rename for display
    result.rename(columns={
        "PN_old": "PN (OLD)",
        "Description_old": "Description (OLD)",
        "bom_qty_old": "Qty (OLD)",
        "PN_new": "PN (NEW)",
        "Description_new": "Description (NEW)",
        "bom_qty_new": "Qty (NEW)"
    }, inplace=True)

    st.subheader("📋 Comparison Result")
    st.dataframe(result, use_container_width=True)

    # ======================
    # EXPORT EXCEL
    # ======================
    output = io.BytesIO()
    result.to_excel(output, index=False)
    output.seek(0)

    wb = load_workbook(output)
    ws = wb.active

    # ======================
    # COLORS
    # ======================
    green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    orange = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    blue = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")

    # Find Status column
    status_col = None
    for col in range(1, ws.max_column + 1):
        if ws.cell(row=1, column=col).value == "Status":
            status_col = col
            break

    # Apply colors
    for row in range(2, ws.max_row + 1):
        status = ws.cell(row=row, column=status_col).value

        if status == "Conforme":
            fill = green
        elif status in ["Missing in BOM 1", "Missing in BOM 2"]:
            fill = red
        elif status == "Qty Difference":
            fill = orange
        else:
            fill = blue

        for col in range(1, ws.max_column + 1):
            ws.cell(row=row, column=col).fill = fill

    # ======================
    # DOWNLOAD
    # ======================
    final_file = io.BytesIO()
    wb.save(final_file)
    final_file.seek(0)

    st.download_button(
        label="📥 Download Excel Result",
        data=final_file,
        file_name="BOM_comparison.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
