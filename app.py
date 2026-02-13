import streamlit as st
import pandas as pd
import os
import io
import re

# 1. Page Configuration
st.set_page_config(page_title="Excel Auto-Filler", layout="wide")
st.title("⚡ Excel Data Filler: Glasses Edition")

# ==========================================
# 🔒 INDESTRUCTIBLE LOADER (LOCKED VERSION)
# ==========================================
@st.cache_data
def load_master():
    current_dir = os.getcwd()
    candidates = [f for f in os.listdir(current_dir) if (f.endswith('.xlsx') or f.endswith('.csv')) and "master_clean" in f and not f.startswith('~$')]
    if not candidates:
        st.error("❌ 'master_clean.xlsx' not found in repository."); st.stop()
    file_path = candidates[0]
    df = None
    try:
        df = pd.read_excel(file_path, dtype=str, engine='openpyxl')
    except Exception:
        strategies = [{'sep': None, 'engine': 'python'}, {'sep': ',', 'engine': 'c'}, {'sep': ';', 'engine': 'c'}, {'sep': '\t', 'engine': 'c'}]
        for enc in ['utf-8', 'cp1252', 'latin1']:
            for strat in strategies:
                try:
                    df = pd.read_csv(file_path, dtype=str, encoding=enc, on_bad_lines='skip', **strat)
                    break
                except: continue
            if df is not None: break
    if df is None:
        st.error(f"❌ Could not read '{file_path}'."); st.stop()
    df.columns = df.columns.astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
    return df

# Load Master Data
raw_master_df = load_master()
target_col = next((c for c in raw_master_df.columns if "items type" in c.lower()), None)
if target_col:
    master_df = raw_master_df[raw_master_df[target_col].str.lower().str.strip() == "glasses"]
    st.success(f"✅ Brain Loaded: {len(master_df)} valid glasses rows.")
else:
    st.error("❌ 'Items type' column missing."); st.stop()

# ==========================================
# 🧠 THE BRAIN: FILLING LOGIC
# ==========================================

def get_col_by_id(df, target_id):
    """Finds a column name in the user's file that contains a specific ID number."""
    for col in df.columns:
        if re.search(f"ID[:\s]+{target_id}\\b", col):
            return col
    return None

def apply_hs_code(row, type_col, mat_col, sport_col):
    """Logic for Rule 1: HS Code"""
    g_type = str(row.get(type_col, '')).strip() if type_col else ""
    material = str(row.get(mat_col, '')).strip().lower() if mat_col else ""
    sport_type = str(row.get(sport_col, '')).strip().lower() if sport_col else ""

    if g_type == "Sunglasses":
        return "90041091", "Type: Sunglasses"
    if g_type == "Frames" and "plastic" in material:
        return "90031100", "Type: Frames + Material: Plastic"
    if g_type == "Frames" and "metal" in material:
        return "90031900", "Type: Frames + Material: Metal"
    if g_type == "Sport glasses":
        if any(x in sport_type for x in ["swimm", "swim", "ski", "snowboard"]):
            return "90049090", f"Type: Sport ({sport_type})"
    return "", "No Match"

def apply_item_description(row, type_col, mat_col):
    """Logic for Rule 2: Item Description"""
    g_type = str(row.get(type_col, '')).strip() if type_col else ""
    material = str(row.get(mat_col, '')).strip().lower() if mat_col else ""

    # Eyeglasses category
    if g_type in ["Frames", "PC Glasses without power", "Driving Glasses without power", "Reading glasses"]:
        return "Eyeglasses", f"Match: {g_type}"
    
    # Sunglasses category
    if g_type == "Sunglasses":
        if "plastic" in material:
            return "Sunglasses, plastic frame", "Sunglasses + Plastic"
        if "metal" in material:
            return "Sunglasses, metal frame", "Sunglasses + Metal"
        return "Sunglasses", "Sunglasses (Unknown Material)"
    
    # Sport glasses
    if g_type == "Sport glasses":
        return "Sport glasses", "Exact match: Sport glasses"
        
    return "", "No Match"

def run_auto_fill(user_df):
    # 1. Identify Columns by ID
    type_col = get_col_by_id(user_df, "13")      # Glasses type
    material_col = get_col_by_id(user_df, "53")  # Main material
    sport_col = get_col_by_id(user_df, "89")     # Sport glasses
    
    # Find Item Description (AP) and HS Code (AO) columns
    # We look for ID: AO/AP or the specific text names
    hs_col = get_col_by_id(user_df, "AO") or "HS Code"
    desc_col = get_col_by_id(user_df, "AP") or "Item description"

    # Ensure columns exist
    if hs_col not in user_df.columns: user_df[hs_col] = ""
    if desc_col not in user_df.columns: user_df[desc_col] = ""

    # 2. Apply Rule 1: HS Code
    hs_results = user_df.apply(lambda row: apply_hs_code(row, type_col, material_col, sport_col), axis=1)
    user_df[hs_col] = [r[0] for r in hs_results]
    hs_reasons = [r[1] for r in hs_results]

    # 3. Apply Rule 2: Item Description
    desc_results = user_df.apply(lambda row: apply_item_description(row, type_col, material_col), axis=1)
    user_df[desc_col] = [r[0] for r in desc_results]
    desc_reasons = [r[1] for r in desc_results]
    
    # 4. Generate Combined Report
    report_df = pd.DataFrame({
        'Type (ID:13)': user_df[type_col] if type_col else "Not Found",
        'HS Code': user_df[hs_col],
        'Item Description': user_df[desc_col],
        'Description Logic': desc_reasons
    })
    
    # Show rows where either rule did something
    modified_rows = report_df[(report_df['HS Code'] != "") | (report_df['Item Description'] != "")]
    
    return user_df, modified_rows

# ==========================================
# 📤 USER INTERFACE
# ==========================================
st.divider()
st.subheader("1. Upload Partial Data")
uploaded_file = st.file_uploader("Choose Excel File", type=['xlsx'])

if uploaded_file:
    user_df = pd.read_excel(uploaded_file, dtype=str)
    st.write(f"Loaded {len(user_df)} rows. Headers preserved.")

    st.divider()
    st.subheader("2. Run Auto-Fill")
    
    if st.button("✨ Auto-Fill Data", type="primary"):
        with st.spinner("Applying Rules 1 & 2..."):
            working_df = user_df.copy()
            filled_df, report = run_auto_fill(working_df)
            
            st.success(f"✅ Processing Complete!")
            
            with st.expander("📊 View Processing Report (Rules 1 & 2)", expanded=True):
                if not report.empty:
                    st.dataframe(report, use_container_width=True)
                else:
                    st.info("No rules were triggered for this file.")
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                filled_df.to_excel(writer, index=False)
            buffer.seek(0)
            
            st.download_button(
                label="📥 Download Formatted Excel",
                data=buffer,
                file_name="filled_glasses_data.xlsx",
                mime="application/vnd.ms-excel"
            )