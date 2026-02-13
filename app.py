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
        # Looks for "ID: 13" or "ID:13" or "ID 13" in the header string
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

def run_auto_fill(user_df):
    # 1. Identify Columns by ID so we don't care about the text name
    type_col = get_col_by_id(user_df, "13")      # Glasses type
    material_col = get_col_by_id(user_df, "53")  # Main material
    sport_col = get_col_by_id(user_df, "89")     # Sport glasses
    hs_col = get_col_by_id(user_df, "AO") or "HS Code" # Check for AO or clean name

    # 2. Add HS Code column if it truly doesn't exist at all
    if hs_col not in user_df.columns:
        user_df[hs_col] = ""

    # 3. Apply Rules
    results = user_df.apply(lambda row: apply_hs_code(row, type_col, material_col, sport_col), axis=1)
    user_df[hs_col] = [r[0] for r in results]
    reasons = [r[1] for r in results]
    
    # 4. Generate Report (Using the ID-found columns for display)
    report_df = pd.DataFrame({
        'Input Type': user_df[type_col] if type_col else "Not Found",
        'Input Material': user_df[material_col] if material_col else "Not Found",
        'Result HS Code': user_df[hs_col],
        'Logic Used': reasons
    })
    filled_only = report_df[report_df['Result HS Code'] != ""]
    
    return user_df, filled_only

# ==========================================
# 📤 USER INTERFACE
# ==========================================
st.divider()
st.subheader("1. Upload Partial Data")
uploaded_file = st.file_uploader("Choose Excel File", type=['xlsx'])

if uploaded_file:
    user_df = pd.read_excel(uploaded_file, dtype=str)
    st.write(f"Loaded {len(user_df)} rows. Found {len(user_df.columns)} columns.")

    st.divider()
    st.subheader("2. Run Auto-Fill")
    
    if st.button("✨ Auto-Fill Data", type="primary"):
        with st.spinner("Applying rules..."):
            # We work on a copy to keep the original safe
            working_df = user_df.copy()
            filled_df, report = run_auto_fill(working_df)
            
            st.success(f"✅ Processing Complete!")
            
            with st.expander("📊 View Processing Report (Testing Mode)", expanded=True):
                if not report.empty:
                    st.dataframe(report, use_container_width=True)
                else:
                    st.info("No HS Codes were filled based on Rule 1 logic.")
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                filled_df.to_excel(writer, index=False)
            buffer.seek(0)
            
            st.download_button(
                label="📥 Download Ready-to-Import File",
                data=buffer,
                file_name="filled_glasses_data.xlsx",
                mime="application/vnd.ms-excel"
            )