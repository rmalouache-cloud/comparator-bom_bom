import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="BOM Comparator", layout="wide")
st.title("📊 BOM Comparison Tool")

old_file = st.file_uploader("📂 Upload OLD BOM", type=["xlsx"])
new_file = st.file_uploader("📂 Upload NEW BOM", type=["xlsx"])

start = st.button("🚀 Start Comparison")

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
    # MERGE EXACT LEVEL (IMPORTANT)
    # =========================
    df = old.merge(
        new,
        on=["PN", "Position"],
        how="outer",
        suffixes=("_old", "_new"),
        indicator=True
    )

    # =========================
    # STATUS
    # =========================
    def get_status(row):

        if row["_merge"] == "left_only":
            return "Missing in BOM2"

        if row["_merge"] == "right_only":
            return "Missing in BOM1"

        if row["bom_qty_old"] != row["bom_qty_new"]:
            return "Qty diff"

        return "Conform"

    df["Status"] = df.apply(get_status, axis=1)

    # =========================
    # FINAL FORMAT EXACT LIKE YOUR SAMPLE
    # =========================
    result = df[[
        "PN",
        "Description_old",
        "bom_qty_old",
        "Position",
        "PN",
        "Description_new",
        "bom_qty_new",
        "Position",
        "Status"
    ]].copy()

    result.columns = [
        "PN",
        "Desc OLD",
        "Qty OLD",
        "Pos OLD",
        "PN NEW",
        "Desc NEW",
        "Qty NEW",
        "Pos NEW",
        "Status"
    ]

    st.dataframe(result, use_container_width=True)
