import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import io

st.set_page_config(page_title="BOM Comparator", layout="wide")

st.title("📊 BOM Comparison Tool")

# Upload files
old_file = st.file_uploader("📂 Upload OLD BOM", type=["xlsx"])
new_file = st.file_uploader("📂 Upload NEW BOM", type=["xlsx"])

if old_file and new_file:

    # Read files
    old_bom = pd.read_excel(old_file)
    new_bom = pd.read_excel(new_file)

    # Clean column names
    old_bom.columns = old_bom.columns.str.strip()
    new_bom.columns = new_bom.columns.str.strip()

    # Required columns
    cols = ["PN", "Description", "bom_qty", "BOM text"]

    # Check columns existence
    for col in cols:
        if col not in old_bom.columns or col not in new_bom.columns:
            st.error(f"❌ Missing column: {col}")
            st.stop()

    # Select columns
    old_bom = old_bom[cols].copy()
    new_bom = new_bom[cols].copy()

    # Rename position column
    old_bom.rename(columns={"BOM text": "Position"}, inplace=True)
    new_bom.rename(columns={"BOM text": "Position"}, inplace=True)

    # Normalize
    for df in [old_bom, new_bom]:
        df["PN"] = df["PN"].astype(str).str.strip().str.upper()
        df["Position"] = df["Position"].astype(str).str.strip().str.upper()

    # Create key
    old_bom["key"] = old_bom["PN"] + "_" + old_bom["Position"]
    new_bom["key"] = new_bom["PN"] + "_" + new_bom["Position"]

    # Merge (IMPORTANT: indicator=True)
    df = old_bom.merge(
        new_bom,
        on="key",
        how="outer",
        suffixes=("_old", "_new"),
        indicator=True
    )

    # Replace NaN to avoid errors
    df.fillna("", inplace=True)

    # Status logic (safe)
    def get_status(row):

        if row["_merge"] == "both":
            if row["bom_qty_old"] == row["bom_qty_new"]:
                return "Conforme"
            else:
                return "Qty Difference"

        elif row["_merge"] == "left_only":
            if row["PN_old"] in new_bom["PN"].values:
                return "Position Difference"
            else:
                return "Missing in BOM 2"

        elif row["_merge"] == "right_only":
            if row["PN_new"] in old_bom["PN"].values:
                return "Position Difference"
            else:
                return "Missing in BOM 1"

        return "Unknown"

    df["Status"] = df.apply(get_status, axis=1)

    # Final table
    result = df[[
        "PN_old", "Description_old", "bom_qty_old", "Position",
        "PN_new", "Description_new", "bom_qty_new", "Position",
        "Status"
    ]]

    st.subheader("📋 Comparison Result")
    st.dataframe(result, use_container_width=True)

    # =========================
    # EXPORT EXCEL WITH COLORS
    # =========================

    output = io.BytesIO()
    result.to_excel(output, index=False)
    output.seek(0)

    wb = load_workbook(output)
    ws = wb.active

    # Colors
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
        elif status == "Position Difference":
            fill = blue
        else:
            continue

        for col in range(1, ws.max_column + 1):
            ws.cell(row=row, column=col).fill = fill

    # Save final file
    final_file = io.BytesIO()
    wb.save(final_file)
    final_file.seek(0)

    # Download button
    st.download_button(
        label="📥 Download Excel Result",
        data=final_file,
        file_name="BOM_comparison.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
