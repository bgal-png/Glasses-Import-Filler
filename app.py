import streamlit as st
import pandas as pd
import os
import io

# 1. Page Configuration
st.set_page_config(page_title="Excel Auto-Filler", layout="wide")
st.title("⚡ Excel Data Filler: Glasses Edition")

# ==========================================
# 🔒 INDESTRUCTIBLE LOADER (LOCKED VERSION)
# 🛑 WARNING: DO NOT MODIFY THIS FUNCTION UNDER ANY CIRCUMSTANCES.
# ==========================================
@st.cache_data
def load_master():
    """
    TRULY INDESTRUCTIBLE LOADER
    1. Tries Excel (.xlsx)
    2. If that fails, tries CSV with Auto-Separator.
    3. If that fails, tries CSV with comma/semicolon explicitly.
    """
    current_dir = os.getcwd()
    candidates = [f for f in os.listdir(current_dir) if (f.endswith('.xlsx') or f.endswith('.csv')) and "master_clean" in f and not f.startswith('~$')]
    
    if not candidates:
        st.error("❌ 'master_clean.xlsx' not found in repository."); st.stop()
    
    file_path = candidates[0]
    df = None
    
    # ATTEMPT 1: EXCEL (Standard)
    try:
        df = pd.read_excel(file_path, dtype=str, engine='openpyxl')
    except Exception:
        # ATTEMPT 2: CSV (Fallback loop)
        strategies = [
            {'sep': None, 'engine': 'python'}, 
            {'sep': ',', 'engine': 'c'},       
            {'sep': ';', 'engine': 'c'},       
            {'sep': '\t', 'engine': 'c'}       
        ]
        
        for enc in ['utf-8', 'cp1252', 'latin1']:
            for strat in strategies:
                try:
                    df = pd.read_csv(file_path, dtype=str, encoding=enc, on_bad_lines='skip', **strat)
                    break
                except: continue
            if df is not None: break
    
    if df is None:
        st.error(f"❌ Could not read '{file_path}'. Tried Excel and all CSV formats.")
        st.stop()

    # Clean headers (Standardize)
    df.columns = df.columns.astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
    return df
# ==========================================
# 🛑 END OF LOCKED LOADER
# ==========================================

# 1. Load the raw data (Safe)
raw_master_df = load_master()

# 2. Apply Filter (Smart Case-Insensitive Search)
target_col = next((c for c in raw_master_df.columns if "items type" in c.lower()), None)
if target_col:
    master_df = raw_master_df[raw_master_df[target_col].str.lower().str.strip() == "glasses"]
    st.success(f"✅ Brain Loaded: {len(master_df)} valid glasses rows.")
else:
    st.error(f"❌ Could not find 'Items type' column. I see these columns: {list(raw_master_df.columns)}")
    st.stop()

# ==========================================
# 🧠 THE BRAIN: FILLING LOGIC
# ==========================================
def apply_hs_code(row):
    """Logic for Rule 1: HS Code"""
    g_type = str(row.get('Glasses type', '')).strip()
    material = str(row.get('Glasses main material', '')).strip().lower()
    sport_type = str(row.get('Sport glasses', '')).strip().lower()

    # 1. Sunglasses
    if g_type == "Sunglasses":
        return "90041091"
    
    # 2. Frames + Plastic
    if g_type == "Frames" and "plastic" in material:
        return "90031100"
    
    # 3. Frames + Metal
    if g_type == "Frames" and "metal" in material:
        return "90031900"
    
    # 4. Sport Glasses (Swim / Ski / Snowboard)
    if g_type == "Sport glasses":
        if "swimm" in sport_type or "swim" in sport_type or "ski" in sport_type or "snowboard" in sport_type:
            return "90049090"
            
    return "" # Default empty if no rule matches

def run_auto_fill(user_df, master_df):
    """
    Takes the user's partial data and standardizes the columns.
    """
    # 1. DEFINE TARGET COLUMNS (33 Standard + 2 New)
    TARGET_COLUMNS = [
        "Glasses type", "Manufacturer", "Brand", "Producing company",
        "Glasses size: glasses width", "Glasses size: temple length", 
        "Glasses size: lens height", "Glasses size: lens width", "Glasses size: bridge",
        "Glasses size: glasses to bend length", "Glasses shape", "Glasses frame type", 
        "Glasses frame color", "Glasses temple color", "Glasses main material", 
        "Glasses lens color", "Glasses lens material", "Glasses lens effect",
        "Sunglasses filter", "UV filter", "SunGlasses RX lenses",
        "Glasses genre", "Glasses usable", "Glasses collection",
        "Items type", "Items packing", "Glasses contain", 
        "Sport glasses", "Glasses frame color effect", "Glasses other features",
        "Glasses clip-on lens color", "Glasses for your face shape", 
        "Glasses lenses no-orders", "Glasses other info",
        "HS Code", "Item description"
    ]
    
    # 2. CREATE MISSING COLUMNS
    for col in TARGET_COLUMNS:
        if col not in user_df.columns:
            user_df[col] = "" 
            
    # 3. APPLY RULES
    # RULE 1: HS Code
    user_df['HS Code'] = user_df.apply(apply_hs_code, axis=1)
    
    # 4. REORDER
    final_cols = TARGET_COLUMNS + [c for c in user_df.columns if c not in TARGET_COLUMNS]
    return user_df[final_cols]

# ==========================================
# 📤 USER INTERFACE
# ==========================================
st.divider()
st.subheader("1. Upload Partial Data")
uploaded_file = st.file_uploader("Choose Excel File", type=['xlsx'])

if uploaded_file:
    try:
        user_df = pd.read_excel(uploaded_file, dtype=str)
    except:
        user_df = pd.read_csv(uploaded_file, dtype=str)
        
    st.write(f"Loaded {len(user_df)} rows.")

    st.divider()
    st.subheader("2. Run Auto-Fill")
    
    if st.button("✨ Auto-Fill Data", type="primary"):
        with st.spinner("Applying rules..."):
            filled_df = run_auto_fill(user_df, master_df)
            st.success("✅ Done! Rule 1 Applied.")
            
            # DOWNLOAD
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