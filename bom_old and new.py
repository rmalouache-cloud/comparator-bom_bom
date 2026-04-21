import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import io

st.set_page_config(page_title="BOM Comparator", layout="wide")
st.title("📊 BOM Comparison Tool")

old_file = st.file_uploader("📂 Upload OLD BOM", type=["xlsx"])
new_file = st.file_uploader("📂 Upload NEW BOM", type=["xlsx"])

# =========================
# SAFE FUNCTION
# =========================
def safe_join(x):
    if not isinstance(x, list):
        return ""
    return ", ".join(str(i) for i in x if pd.notna(i))

def extract_components_by_type(df, bom_type):
    """
    Extrait les composants CKD ou SKD selon la ligne de démarcation
    CKD: de "ASS'Y - MAIN BOARD（CKD）" jusqu'à "BARCODE LABEL"
    SKD: tout le reste
    """
    if bom_type == "CKD":
        # Trouver l'index de début (ASS'Y - MAIN BOARD（CKD）)
        start_idx = None
        for idx, desc in enumerate(df['Description']):
            if 'ASS\'Y - MAIN BOARD（CKD）' in str(desc).upper() or 'ASSY - MAIN BOARD（CKD）' in str(desc).upper():
                start_idx = idx
                break
        
        # Trouver l'index de fin (BARCODE LABEL)
        end_idx = None
        for idx, desc in enumerate(df['Description']):
            if 'BARCODE LABEL' in str(desc).upper():
                end_idx = idx
                break
        
        if start_idx is not None and end_idx is not None:
            return df.iloc[start_idx:end_idx+1].copy()
        elif start_idx is not None:
            return df.iloc[start_idx:].copy()
        else:
            return pd.DataFrame()  # Pas de composants CKD trouvés
    else:  # SKD
        # Trouver l'index de début (ASS'Y - MAIN BOARD（CKD）)
        start_idx = None
        for idx, desc in enumerate(df['Description']):
            if 'ASS\'Y - MAIN BOARD（CKD）' in str(desc).upper() or 'ASSY - MAIN BOARD（CKD）' in str(desc).upper():
                start_idx = idx
                break
        
        # Trouver l'index de fin (BARCODE LABEL)
        end_idx = None
        for idx, desc in enumerate(df['Description']):
            if 'BARCODE LABEL' in str(desc).upper():
                end_idx = idx
                break
        
        if start_idx is not None and end_idx is not None:
            # SKD = tout avant CKD + tout après BARCODE LABEL
            before_ckd = df.iloc[:start_idx].copy()
            after_ckd = df.iloc[end_idx+1:].copy()
            return pd.concat([before_ckd, after_ckd], ignore_index=True)
        elif start_idx is not None:
            return df.iloc[:start_idx].copy()
        else:
            return df.copy()  # Pas de ligne CKD trouvée, tout est SKD

def run_comparison(old_df, new_df, component_type="GENERAL"):
    """Exécute la comparaison entre deux DataFrames"""
    
    cols = ["PN", "Description", "bom_qty", "BOM text"]
    
    old = old_df[cols].copy()
    new = new_df[cols].copy()
    
    old.rename(columns={"BOM text": "Position"}, inplace=True)
    new.rename(columns={"BOM text": "Position"}, inplace=True)
    
    # Clean data
    for df in [old, new]:
        df["PN"] = df["PN"].astype(str).str.strip().str.upper()
        df["Description"] = df["Description"].astype(str).str.strip().str.upper()
        df["Position"] = df["Position"].astype(str).str.strip().str.upper()
        
        df["bom_qty"] = (
            df["bom_qty"]
            .astype(str)
            .str.replace(",", ".", regex=False)
        )
        df["bom_qty"] = pd.to_numeric(df["bom_qty"], errors="coerce").fillna(0)
    
    # Group by PN only
    old = old.groupby(["PN"], as_index=False).agg({
        "Description": "first",
        "bom_qty": "sum",
        "Position": list
    })
    
    new = new.groupby(["PN"], as_index=False).agg({
        "Description": "first",
        "bom_qty": "sum",
        "Position": list
    })
    
    # Merge on PN
    df = old.merge(
        new,
        on="PN",
        how="outer",
        suffixes=("_old", "_new"),
        indicator=True
    )
    
    # Logic
    def get_status(row):
        if row["_merge"] == "left_only":
            return "Missing in BOM2"
        
        if row["_merge"] == "right_only":
            return "Missing in BOM1"
        
        qty_old = row["bom_qty_old"]
        qty_new = row["bom_qty_new"]
        
        pos_old = set(row["Position_old"]) if isinstance(row["Position_old"], list) else set()
        pos_new = set(row["Position_new"]) if isinstance(row["Position_new"], list) else set()
        
        if qty_old != qty_new:
            return "Qty diff"
        
        if pos_old != pos_new:
            return "Position diff"
        
        return "Conform"
    
    df["Status"] = df.apply(get_status, axis=1)
    
    # Build result
    result = []
    
    for _, row in df.iterrows():
        pos_old = row["Position_old"] if isinstance(row["Position_old"], list) else []
        pos_new = row["Position_new"] if isinstance(row["Position_new"], list) else []
        
        result.append({
            "Type": component_type,
            "PN": row["PN"] if row["_merge"] != "right_only" else "",
            "Desc OLD": row.get("Description_old", ""),
            "Qty OLD": row.get("bom_qty_old", 0),
            "Pos OLD": safe_join(pos_old),
            "PN NEW": row["PN"] if row["_merge"] != "left_only" else "",
            "Desc NEW": row.get("Description_new", ""),
            "Qty NEW": row.get("bom_qty_new", 0),
            "Pos NEW": safe_join(pos_new),
            "Status": row["Status"]
        })
    
    return pd.DataFrame(result)

# =========================
# 3 BOUTONS POUR LA COMPARAISON
# =========================
st.markdown("---")
st.subheader("🔧 Sélectionnez le type de comparaison")

col1, col2, col3 = st.columns(3)

with col1:
    ckd_button = st.button("📟 CKD", use_container_width=True)
with col2:
    skd_button = st.button("🔌 SKD", use_container_width=True)
with col3:
    both_button = st.button("🔄 CKD/SKD", use_container_width=True)

if ckd_button or skd_button or both_button:
    
    if old_file is None or new_file is None:
        st.error("❌ Upload both files")
        st.stop()
    
    # Read files
    old = pd.read_excel(old_file)
    new = pd.read_excel(new_file)
    
    old.columns = old.columns.str.strip()
    new.columns = new.columns.str.strip()
    
    all_results = []
    
    # =========================
    # CAS 1: BOUTON CKD SEULEMENT
    # =========================
    if ckd_button:
        bom_type = "CKD"
        st.info("📟 Comparaison des composants CKD uniquement")
        
        old_filtered = extract_components_by_type(old, "CKD")
        new_filtered = extract_components_by_type(new, "CKD")
        
        if old_filtered.empty and new_filtered.empty:
            st.warning("⚠️ Aucun composant CKD trouvé dans les fichiers")
        else:
            st.info(f"📊 Composants CKD trouvés: {len(old_filtered)} dans OLD BOM, {len(new_filtered)} dans NEW BOM")
            result = run_comparison(old_filtered, new_filtered, "CKD")
            all_results.append(result)
    
    # =========================
    # CAS 2: BOUTON SKD SEULEMENT
    # =========================
    elif skd_button:
        bom_type = "SKD"
        st.info("🔌 Comparaison des composants SKD uniquement")
        
        old_filtered = extract_components_by_type(old, "SKD")
        new_filtered = extract_components_by_type(new, "SKD")
        
        if old_filtered.empty and new_filtered.empty:
            st.warning("⚠️ Aucun composant SKD trouvé dans les fichiers")
        else:
            st.info(f"📊 Composants SKD trouvés: {len(old_filtered)} dans OLD BOM, {len(new_filtered)} dans NEW BOM")
            result = run_comparison(old_filtered, new_filtered, "SKD")
            all_results.append(result)
    
    # =========================
    # CAS 3: BOUTON CKD/SKD (LES DEUX ENSEMBLE)
    # =========================
    elif both_button:
        st.info("🔄 Comparaison CKD + SKD ensemble (tous les composants)")
        
        # Extraire CKD et SKD séparément
        old_ckd = extract_components_by_type(old, "CKD")
        old_skd = extract_components_by_type(old, "SKD")
        old_full = pd.concat([old_ckd, old_skd], ignore_index=True)
        
        new_ckd = extract_components_by_type(new, "CKD")
        new_skd = extract_components_by_type(new, "SKD")
        new_full = pd.concat([new_ckd, new_skd], ignore_index=True)
        
        st.info(f"📊 Comparaison complète: {len(old_full)} composants dans OLD BOM, {len(new_full)} dans NEW BOM")
        st.info(f"   - CKD: {len(old_ckd)} (OLD) / {len(new_ckd)} (NEW)")
        st.info(f"   - SKD: {len(old_skd)} (OLD) / {len(new_skd)} (NEW)")
        
        result = run_comparison(old_full, new_full, "COMPLET")
        all_results.append(result)
    
    # =========================
    # AFFICHER LES RÉSULTATS
    # =========================
    if all_results:
        final_result = pd.concat(all_results, ignore_index=True)
        
        # Fix Streamlit crash
        for col in final_result.columns:
            final_result[col] = final_result[col].astype(str)
        
        st.success(f"✅ Comparaison terminée! {len(final_result)} lignes trouvées")
        st.dataframe(final_result, use_container_width=True)
        
        # Export Excel with colors
        output = io.BytesIO()
        final_result.to_excel(output, index=False)
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
                break
        
        if status_col:
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
        
        # Nom du fichier selon le type de comparaison
        if ckd_button:
            filename = "BOM_comparison_CKD.xlsx"
        elif skd_button:
            filename = "BOM_comparison_SKD.xlsx"
        else:
            filename = "BOM_comparison_CKD_SKD.xlsx"
        
        st.download_button(
            f"📥 Télécharger Excel",
            final_file,
            filename
        )
    else:
        st.error("❌ Aucun résultat à afficher")
