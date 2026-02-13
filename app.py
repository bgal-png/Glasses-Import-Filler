import streamlit as st
import pandas as pd
import os
import io

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
def apply_hs_code(row):
    """Logic for Rule 1: HS Code"""
    g_type = str(row.get('Glasses type', '')).strip()
    material = str(row.get('Glasses main material', '')).strip().lower()
    sport_type = str(row.get('Sport glasses', '')).strip().lower()

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
    """
    Standardizes the file WITHOUT changing the original column order.
    """
    # All 35 columns we care about
    REQUIRED_COLUMNS = [
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
    
    # 1. ADD MISSING COLUMNS (Only if they don't exist in the uploaded file)
    for col in REQUIRED_COLUMNS:
        if col not in user_df.columns:
            user_df[col] = "" 
            
    # 2. APPLY RULES (Directly onto existing columns)
    results = user_df.apply(apply_hs_code, axis=1)
    user_df['HS Code'] = [r[0] for r in results]
    reasons = [r[1] for r in results]
    
    # 3. GENERATE REPORT (For Testing)
    report_df = user_df[['Glasses type', 'Glasses main material', 'HS Code']].copy()
    report_df['Reason/Logic'] = reasons
    filled_only = report_df[report_df['HS Code'] != ""]
    
    # 4. RETURN (User's order is preserved naturally)
    return user_df, filled_only

# ==========================================
# 📤 USER INTERFACE
# ==========================================
st.divider()
st.subheader("1. Upload Partial Data")
uploaded_file = st.file_uploader("Choose Excel File", type=['xlsx'])

if uploaded_file:
    user_df = pd.read_excel(uploaded_file, dtype=str)
    st.write(f"Loaded {len(user_df)} rows with columns: {list(user_df.columns)}")

    st.divider()
    st.subheader("2. Run Auto-Fill")
    
    if st.button("✨ Auto-Fill Data", type="primary"):
        with st.spinner("Applying rules..."):
            filled_df, report = run_auto_fill(user_df)
            
            st.success(f"✅ Processing Complete!")
            
            # --- TESTING REPORT SECTION ---
            with st.expander("📊 View Processing Report (Testing Mode)", expanded=True):
                if not report.empty:
                    st.dataframe(report, use_container_width=True)
                else:
                    st.info("No HS Codes were filled based on Rule 1.")
            
            # DOWNLOAD
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