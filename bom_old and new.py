import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.title("BOM Comparator SMT")

old_file = st.file_uploader("Upload OLD BOM", type=["xlsx"])
new_file = st.file_uploader("Upload NEW BOM", type=["xlsx"])

if st.button("Start Comparison"):

    old = pd.read_excel(old_file)
    new = pd.read_excel(new_file)

    # Clean
    old.columns = old.columns.str.strip()
    new.columns = new.columns.str.strip()

    cols = ["PN", "Description", "bom_qty", "BOM text"]
    old = old[cols].copy()
    new = new[cols].copy()

    old.rename(columns={"BOM text": "Position"}, inplace=True)
    new.rename(columns={"BOM text": "Position"}, inplace=True)

    for df in [old, new]:
        df["PN"] = df["PN"].astype(str).str.strip().str.upper()
        df["Description"] = df["Description"].astype(str).str.strip().str.upper()
        df["Position"] = df["Position"].astype(str).str.strip().str.upper()

    # ======================
    # GROUP BY PN 🔥
    # ======================
    old_group = old.groupby("PN").agg({
        "Description": "first",
        "Position": lambda x: list(x)
    }).reset_index()

    new_group = new.groupby("PN").agg({
        "Description": "first",
        "Position": lambda x: list(x)
    }).reset_index()

    # Convert to set
    old_group["Position"] = old_group["Position"].apply(set)
    new_group["Position"] = new_group["Position"].apply(set)

    # Qty = nombre de positions
    old_group["Qty"] = old_group["Position"].apply(len)
    new_group["Qty"] = new_group["Position"].apply(len)

    # ======================
    # MERGE PN
    # ======================
    df = old_group.merge(
        new_group,
        on="PN",
        how="outer",
        suffixes=("_old", "_new"),
        indicator=True
    )

    # ======================
    # STATUS LOGIC
    # ======================
    def get_status(row):

        if row["_merge"] == "both":

            if row["Description_old"] != row["Description_new"]:
                return "Check Description"

            if row["Position_old"] == row["Position_new"]:
                return "Conform"

            elif row["Qty_old"] != row["Qty_new"]:
                return "Qty Difference"

            else:
                return "Position Difference"

        elif row["_merge"] == "left_only":
            return "Missing in BOM2"

        elif row["_merge"] == "right_only":
            return "Missing in BOM1"

    df["Status"] = df.apply(get_status, axis=1)

    # Convert set to string for display
    df["Position_old"] = df["Position_old"].astype(str)
    df["Position_new"] = df["Position_new"].astype(str)

    st.dataframe(df)

    # ======================
    # EXPORT EXCEL
    # ======================
    output = io.BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    wb = load_workbook(output)
    ws = wb.active

    green = PatternFill(start_color="C6EFCE", fill_type="solid")
    red = PatternFill(start_color="FFC7CE", fill_type="solid")
    orange = PatternFill(start_color="FFEB9C", fill_type="solid")
    blue = PatternFill(start_color="BDD7EE", fill_type="solid")

    for col in range(1, ws.max_column + 1):
        if ws.cell(1, col).value == "Status":
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
        else:
            fill = blue

        for col in range(1, ws.max_column + 1):
            ws.cell(row, col).fill = fill

    final = io.BytesIO()
    wb.save(final)
    final.seek(0)

    st.download_button("Download Excel", final, "result.xlsx")
