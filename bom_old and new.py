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

    for col in cols:
        if col not in old.columns or col not in new.columns:
            st.error(f"❌ Missing column: {col}")
            st.stop()

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

        df["bom_qty"] = (
            df["bom_qty"]
            .astype(str)
            .str.replace(",", ".", regex=False)
        )
        df["bom_qty"] = pd.to_numeric(df["bom_qty"], errors="coerce")

    # ======================
    # REMOVE DUPLICATES
    # ======================
    old = old.groupby(["PN", "Position", "Description"], as_index=False)["bom_qty"].sum()
    new = new.groupby(["PN", "Position", "Description"], as_index=False)["bom_qty"].sum()

    # ======================
    # MERGE
    # ======================
    df = old.merge(
        new,
        on=["PN", "Position"],
        how="outer",
        suffixes=("_old", "_new"),
        indicator=True
    )

    df = df.fillna("")

    df["bom_qty_old"] = pd.to_numeric(df["bom_qty_old"], errors="coerce").fillna(0)
    df["bom_qty_new"] = pd.to_numeric(df["bom_qty_new"], errors="coerce").fillna(0)

    # ======================
    # STATUS
    # ======================
    def get_status(row):

        if row["_merge"] == "both":

            if (
                row["Description_old"] == row["Description_new"] and
                row["bom_qty_old"] == row["bom_qty_new"]
            ):
                return "Conform"

            elif row["bom_qty_old"] != row["bom_qty_new"]:
                return "Qty Difference"

            else:
                return "Position Difference"

        elif row["_merge"] == "left_only":
            return "Missing in BOM2"

        elif row["_merge"] == "right_only":
            return "Missing in BOM1"

        return "Unknown"

    df["Status"] = df.apply(get_status, axis=1)

    # ======================
    # FORMAT FINAL OUTPUT (IMPORTANT 🔥)
    # ======================
    def format_row(row):

        if row["Status"] == "Missing in BOM2":
            return pd.Series({
                "PN": row["PN"],
                "Desc OLD": row["Description_old"],
                "Qty OLD": row["bom_qty_old"],
                "Pos OLD": row["Position"],

                "PN NEW": "",
                "Desc NEW": "",
                "Qty NEW": "",
                "Pos NEW": "",

                "Status": "Missing in BOM2"
            })

        elif row["Status"] == "Missing in BOM1":
            return pd.Series({
                "PN": "",
                "Desc OLD": "",
                "Qty OLD": "",
                "Pos OLD": "",

                "PN NEW": row["PN"],
                "Desc NEW": row["Description_new"],
                "Qty NEW": row["bom_qty_new"],
                "Pos NEW": row["Position"],

                "Status": "Missing in BOM1"
            })

        else:
            return pd.Series({
                "PN": row["PN"],
                "Desc OLD": row["Description_old"],
                "Qty OLD": row["bom_qty_old"],
                "Pos OLD": row["Position"],

                "PN NEW": row["PN"],
                "Desc NEW": row["Description_new"],
                "Qty NEW": row["bom_qty_new"],
                "Pos NEW": row["Position"],

                "Status": row["Status"]
            })

    result = df.apply(format_row, axis=1)

    st.dataframe(result, use_container_width=True)

    # ======================
    # EXPORT EXCEL
    # ======================
    output = io.BytesIO()
    result.to_excel(output, index=False)
    output.seek(0)

    wb = load_workbook(output)
    ws = wb.active

    colors = {
        "Conform": "C6EFCE",
        "Missing": "FFC7CE",
        "Qty Difference": "FFEB9C",
        "Position Difference": "BDD7EE"
    }

    status_col = None
    for col in range(1, ws.max_column + 1):
        if ws.cell(row=1, column=col).value == "Status":
            status_col = col

    for row in range(2, ws.max_row + 1):
        status = ws.cell(row=row, column=status_col).value

        for key in colors:
            if key in str(status):
                fill = PatternFill(start_color=colors[key], fill_type="solid")
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
