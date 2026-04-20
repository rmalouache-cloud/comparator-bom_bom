import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import io

st.set_page_config(page_title="BOM Comparator", layout="wide")
st.title("📊 BOM Comparison Tool")

old_file = st.file_uploader("📂 Upload OLD BOM", type=["xlsx"])
new_file = st.file_uploader("📂 Upload NEW BOM", type=["xlsx"])

start = st.button("🚀 Start Comparison")

if start:

    if old_file is None or new_file is None:
        st.error("❌ Upload both files")
        st.stop()

    # ======================
    # READ FILES
    # ======================
    old = pd.read_excel(old_file)
    new = pd.read_excel(new_file)

    old.columns = old.columns.str.strip()
    new.columns = new.columns.str.strip()

    cols = ["PN", "Description", "bom_qty", "BOM text"]

    old = old[cols].copy()
    new = new[cols].copy()

    old.rename(columns={"BOM text": "Position"}, inplace=True)
    new.rename(columns={"BOM text": "Position"}, inplace=True)

    # ======================
    # CLEAN
    # ======================
    for df in [old, new]:
        df["PN"] = df["PN"].astype(str).str.strip().str.upper()
        df["Description"] = df["Description"].astype(str).str.strip().str.upper()
        df["Position"] = df["Position"].astype(str).str.strip().str.upper()
        df["bom_qty"] = pd.to_numeric(df["bom_qty"], errors="coerce")

    # ======================
    # MERGE ON PN ONLY 🔥
    # ======================
    df = old.merge(
        new,
        on="PN",
        how="outer",
        suffixes=("_old", "_new"),
        indicator=True
    )

    # Fill text only
    text_cols = df.select_dtypes(include=["object"]).columns
    df[text_cols] = df[text_cols].fillna("")

    # ======================
    # STATUS LOGIC 🔥
    # ======================
    def get_status(row):

        if row["_merge"] == "both":

            if (
                row["Description_old"] == row["Description_new"] and
                row["bom_qty_old"] == row["bom_qty_new"] and
                row["Position_old"] == row["Position_new"]
            ):
                return "Conform"

            elif (
                row["Description_old"] == row["Description_new"] and
                row["bom_qty_old"] != row["bom_qty_new"]
            ):
                return "Qty Difference"

            elif (
                row["Description_old"] == row["Description_new"] and
                row["bom_qty_old"] == row["bom_qty_new"] and
                row["Position_old"] != row["Position_new"]
            ):
                return "Position Difference"

            else:
                return "Check Manually"

        elif row["_merge"] == "right_only":
            return "Missing in BOM1"

        elif row["_merge"] == "left_only":
            return "Missing in BOM2"

        return "Unknown"

    df["Status"] = df.apply(get_status, axis=1)

    # ======================
    # FINAL TABLE
    # ======================
    result = df[[
        "PN",
        "Description_old", "bom_qty_old", "Position_old",
        "Description_new", "bom_qty_new", "Position_new",
        "Status"
    ]]

    result.rename(columns={
        "Description_old": "Desc OLD",
        "bom_qty_old": "Qty OLD",
        "Position_old": "Pos OLD",
        "Description_new": "Desc NEW",
        "bom_qty_new": "Qty NEW",
        "Position_new": "Pos NEW"
    }, inplace=True)

    st.dataframe(result, use_container_width=True)

    # ======================
    # EXPORT EXCEL
    # ======================
    output = io.BytesIO()
    result.to_excel(output, index=False)
    output.seek(0)

    wb = load_workbook(output)
    ws = wb.active

    # COLORS
    green = PatternFill(start_color="C6EFCE", fill_type="solid")
    red = PatternFill(start_color="FFC7CE", fill_type="solid")
    orange = PatternFill(start_color="FFEB9C", fill_type="solid")
    blue = PatternFill(start_color="BDD7EE", fill_type="solid")

    # Find Status column
    for col in range(1, ws.max_column + 1):
        if ws.cell(row=1, column=col).value == "Status":
            status_col = col
            break

    # Apply colors
    for row in range(2, ws.max_row + 1):
        status = ws.cell(row=row, column=status_col).value

        if status == "Conform":
            fill = green
        elif "Missing" in status:
            fill = red
        elif status == "Qty Difference":
            fill = orange
        elif status == "Position Difference":
            fill = blue
        else:
            continue

        for col in range(1, ws.max_column + 1):
            ws.cell(row=row, column=col).fill = fill

    final_file = io.BytesIO()
    wb.save(final_file)
    final_file.seek(0)

    st.download_button(
        "📥 Download Excel",
        final_file,
        "BOM_comparison.xlsx"
    )
