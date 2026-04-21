import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="BOM Comparator", layout="wide")
st.title("📊 BOM Comparison Tool")

old_file = st.file_uploader("📂 Upload OLD BOM", type=["xlsx"])
new_file = st.file_uploader("📂 Upload NEW BOM", type=["xlsx"])

start = st.button("🚀 Start Comparison")

# =========================
def safe_join(x):
    if isinstance(x, list):
        return ", ".join(str(i) for i in x)
    return str(x)

# =========================
if start:

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

        df["bom_qty"] = pd.to_numeric(df["bom_qty"], errors="coerce").fillna(0)

        df = df[df["PN"].notna()]
        df = df[df["PN"] != ""]
        df = df[df["PN"].str.lower() != "nan"]

    # =========================
    # GROUP BY PN (NO SUM, KEEP LISTS)
    # =========================
    old_g = old.groupby("PN").agg({
        "Description": "first",
        "bom_qty": "first",
        "Position": list
    }).reset_index()

    new_g = new.groupby("PN").agg({
        "Description": "first",
        "bom_qty": "first",
        "Position": list
    }).reset_index()

    # =========================
    # MERGE PN ONLY
    # =========================
    df = old_g.merge(
        new_g,
        on="PN",
        how="outer",
        suffixes=("_old", "_new"),
        indicator=True
    )

    # =========================
    # LOGIC FIXED
    # =========================
    def get_status(row):

        if row["_merge"] == "left_only":
            return "Missing in BOM2"

        if row["_merge"] == "right_only":
            return "Missing in BOM1"

        qty_old = row["bom_qty_old"]
        qty_new = row["bom_qty_new"]

        pos_old = set(row["Position_old"] if isinstance(row["Position_old"], list) else [])
        pos_new = set(row["Position_new"] if isinstance(row["Position_new"], list) else [])

        # ✔ PRIORITY 1: qty diff
        if qty_old != qty_new:
            return "Qty diff"

        # ✔ PRIORITY 2: position diff
        if pos_old != pos_new:
            return "Position diff"

        return "Conform"

    df["Status"] = df.apply(get_status, axis=1)

    # =========================
    # FINAL OUTPUT CLEAN
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
