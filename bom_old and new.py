import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.title("BOM Comparator - SMT Level")

old_file = st.file_uploader("Upload OLD BOM", type=["xlsx"])
new_file = st.file_uploader("Upload NEW BOM", type=["xlsx"])

if st.button("Start Comparison"):

    old = pd.read_excel(old_file)
    new = pd.read_excel(new_file)

    # Clean columns
    old.columns = old.columns.str.strip()
    new.columns = new.columns.str.strip()

    cols = ["PN", "Description", "bom_qty", "BOM text"]

    old = old[cols].copy()
    new = new[cols].copy()

    old.rename(columns={"BOM text": "Position"}, inplace=True)
    new.rename(columns={"BOM text": "Position"}, inplace=True)

    # Normalize
    for df in [old, new]:
        df["PN"] = df["PN"].astype(str).str.strip().str.upper()
        df["Description"] = df["Description"].astype(str).str.strip()
        df["Position"] = df["Position"].astype(str).str.strip().str.upper()
        df["bom_qty"] = pd.to_numeric(df["bom_qty"], errors="coerce")

    # =========================
    # MERGE PN + POSITION 🔥
    # =========================
    df = old.merge(
        new,
        on=["PN", "Position"],
        how="outer",
        suffixes=("_old", "_new"),
        indicator=True
    )

    # =========================
    # STATUS LOGIC
    # =========================
    def get_status(row):

        if row["_merge"] == "both":

            if row["bom_qty_old"] == row["bom_qty_new"]:
                return "Conform"
            else:
                return "Qty Difference"

        elif row["_merge"] == "left_only":
            return "Missing in BOM2"

        elif row["_merge"] == "right_only":
            return "Missing in BOM1"

    df["Remarks"] = df.apply(get_status, axis=1)

    # =========================
    # HANDLE POSITION DIFFERENCE 🔥
    # =========================
    # détecter même PN mais position différente
    for i, row in df.iterrows():

        if row["Remarks"] in ["Missing in BOM1", "Missing in BOM2"]:

            # chercher même PN dans autre BOM
            if row["PN"] in old["PN"].values and row["PN"] in new["PN"].values:
                df.at[i, "Remarks"] = "Position Difference"

    # =========================
    # FINAL FORMAT
    # =========================
    result = df[[
        "PN", "Description_old", "bom_qty_old", "Position",
        "Description_new", "bom_qty_new", "Position",
        "Remarks"
    ]]

    result.rename(columns={
        "PN": "PN",
        "Description_old": "Desc OLD",
        "bom_qty_old": "Qty OLD",
        "Description_new": "Desc NEW",
        "bom_qty_new": "Qty NEW"
    }, inplace=True)

    st.dataframe(result, use_container_width=True)

    # =========================
    # EXPORT EXCEL
    # =========================
    output = io.BytesIO()
    result.to_excel(output, index=False)
    output.seek(0)

    wb = load_workbook(output)
    ws = wb.active

    green = PatternFill(start_color="C6EFCE", fill_type="solid")
    red = PatternFill(start_color="FFC7CE", fill_type="solid")
    orange = PatternFill(start_color="FFEB9C", fill_type="solid")
    blue = PatternFill(start_color="BDD7EE", fill_type="solid")

    for col in range(1, ws.max_column + 1):
        if ws.cell(1, col).value == "Remarks":
            status_col = col
            break

    for row in range(2, ws.max_row + 1):
        status = ws.cell(row, status_col).value

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
            ws.cell(row, col).fill = fill

    final = io.BytesIO()
    wb.save(final)
    final.seek(0)

    st.download_button("Download Excel", final, "BOM_result.xlsx")
