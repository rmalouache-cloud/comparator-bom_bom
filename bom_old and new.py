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

def safe_join(x):
    if isinstance(x, list):
        return ", ".join(str(i) for i in x if pd.notna(i))
    return str(x)

if start:

    if old_file is None or new_file is None:
        st.error("❌ Upload both files")
        st.stop()

    # =========================
    # READ
    # =========================
    old = pd.read_excel(old_file)
    new = pd.read_excel(new_file)

    old.columns = old.columns.str.strip()
    new.columns = new.columns.str.strip()

    cols = ["PN", "Description", "bom_qty", "BOM text"]

    old = old[cols].copy()
    new = new[cols].copy()

    old.rename(columns={"BOM text": "Position"}, inplace=True)
    new.rename(columns={"BOM text": "Position"}, inplace=True)

    # =========================
    # CLEAN
    # =========================
    for df in [old, new]:
        df["PN"] = df["PN"].astype(str).str.strip().str.upper()
        df["Description"] = df["Description"].astype(str).str.strip().str.upper()
        df["Position"] = df["Position"].astype(str).str.strip().str.upper()

        df["bom_qty"] = pd.to_numeric(
            df["bom_qty"].astype(str).str.replace(",", ".", regex=False),
            errors="coerce"
        ).fillna(0)

        df = df[df["PN"].notna()]
        df = df[df["PN"] != ""]
        df = df[df["PN"].str.lower() != "nan"]

    # =========================
    # GROUP (NO SUM !!)
    # =========================
    old = old.groupby("PN", as_index=False).agg({
        "Description": "first",
        "bom_qty": "first",        # ❌ PAS DE SUM
        "Position": lambda x: list(x)
    })

    new = new.groupby("PN", as_index=False).agg({
        "Description": "first",
        "bom_qty": "first",        # ❌ PAS DE SUM
        "Position": lambda x: list(x)
    })

    # =========================
    # MERGE CLEAN
    # =========================
    df = old.merge(
        new,
        on="PN",
        how="outer",
        suffixes=("_old", "_new"),
        indicator=True
    )

    # =========================
    # STATUS LOGIC
    # =========================
    def get_status(row):

        if row["_merge"] == "left_only":
            return "Missing in BOM2"

        if row["_merge"] == "right_only":
            return "Missing in BOM1"

        if row["bom_qty_old"] != row["bom_qty_new"]:
            return "Qty diff"

        pos_old = set(row["Position_old"]) if isinstance(row["Position_old"], list) else set()
        pos_new = set(row["Position_new"]) if isinstance(row["Position_new"], list) else set()

        if pos_old != pos_new:
            return "Position diff"

        return "Conform"

    df["Status"] = df.apply(get_status, axis=1)

    # =========================
    # FINAL FORMAT (PROPRE TABLE)
    # =========================
    result = pd.DataFrame({
        "PN": df["PN"],

        "Desc OLD": df["Description_old"],
        "Qty OLD": df["bom_qty_old"],
        "Pos OLD": df["Position_old"].apply(safe_join),

        "PN NEW": df["PN"],
        "Desc NEW": df["Description_new"],
        "Qty NEW": df["bom_qty_new"],
        "Pos NEW": df["Position_new"].apply(safe_join),

        "Status": df["Status"]
    })

    st.dataframe(result, use_container_width=True)

    # =========================
    # EXPORT EXCEL
    # =========================
    output = io.BytesIO()
    result.to_excel(output, index=False)
    output.seek(0)

    wb = load_workbook(output)
    ws = wb.active

    colors = {
        "Conform": "C6EFCE",
        "Missing in BOM1": "FFC7CE",
        "Missing in BOM2": "FFC7CE",
        "Qty diff": "FFEB9C",
        "Position diff": "BDD7EE"
    }

    status_col = None
    for col in range(1, ws.max_column + 1):
        if ws.cell(row=1, column=col).value == "Status":
            status_col = col

    for row in range(2, ws.max_row + 1):
        status = ws.cell(row=row, column=status_col).value

        for key, color in colors.items():
            if status == key:
                fill = PatternFill(start_color=color, fill_type="solid")
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
