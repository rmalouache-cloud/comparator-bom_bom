# =========================
# MERGE CORRECT
# =========================
df = old.merge(
    new,
    on="PN",
    how="outer",
    suffixes=("_old", "_new"),
    indicator=True
)

# =========================
# STATUS LOGIC FIXED
# =========================
def get_status(row):

    if row["_merge"] == "left_only":
        return "Missing in BOM2"

    if row["_merge"] == "right_only":
        return "Missing in BOM1"

    qty_old = row["bom_qty_old"]
    qty_new = row["bom_qty_new"]

    pos_old = set(str(row["Position_old"]).split(",")) if pd.notna(row["Position_old"]) else set()
    pos_new = set(str(row["Position_new"]).split(",")) if pd.notna(row["Position_new"]) else set()

    if qty_old != qty_new:
        return "Qty diff"

    if pos_old != pos_new:
        return "Position diff"

    return "Conform"
