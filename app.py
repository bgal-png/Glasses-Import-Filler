import streamlit as st
import pandas as pd
import os
import io

# 1. Page Configuration
st.set_page_config(page_title="Excel Auto-Filler", layout="wide")
st.title("⚡ Excel Data Filler: Glasses Edition")

# ==========================================
# 🔒 LOADER: Indestructible + Filtered
# ==========================================
@st.cache_data
def load_master():
    """
    1. Loads master_clean.xlsx (Indestructible Mode)
    2. Filters ONLY for 'Glasses' rows (Just like Validator)
    """
    current_dir = os.getcwd()
    candidates = [f for f in os.listdir(current_dir) if (f.endswith('.xlsx') or f.endswith('.csv')) and "master_clean" in f and not f.startswith('~$')]
    
    if not candidates:
        st.error("❌ 'master_clean.xlsx' not found in repository."); st.stop()
    
    file_path = candidates[0]
    df = None
    
    # ATTEMPT 1: EXCEL
    try:
        df = pd.read_excel(file_path, dtype=str, engine='openpyxl')
    except Exception:
        # ATTEMPT 2: CSV
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
        st.error(f"❌ Could not read '{file_path}'."); st.stop()

    # Clean headers
    df.columns = df.columns.astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
    
    # 🔍 FILTER LOGIC (The new part)
    target_col = next((c for c in df.columns if "Items type" in c), None)
    if target_col:
        # Only keep rows where Items type is "Glasses"
        df = df[df[target_col] == "Glasses"]
        return df
    else:
        st.error("❌ 'Items type' column missing in Master File."); st.stop()

# Load Master immediately
master_df = load_master()
if not master_df.empty:
    st.success(f"✅ Brain Loaded: {len(master_df)} valid glasses rows.")

# ==========================================
# 🧠 THE BRAIN: FILLING LOGIC
# ==========================================
def run_auto_fill(user_df, master_df):
    """
    Takes the user's partial data and fills in the blanks.
    """
    # 1. DEFINE TARGET COLUMNS (The Standard 33)
    TARGET_COLUMNS = [
        "Glasses type", "Manufacturer", "Brand", "Producing company",
        "Glasses size: glasses width", "Glasses size: temple length", 
        "Glasses size: lens height", "Glasses size: lens width", "Glasses size: bridge",
        "Glasses shape", "Glasses frame type", "Glasses frame color", 
        "Glasses temple color", "Glasses main material", 
        "Glasses lens color", "Glasses lens material", "Glasses lens effect",
        "Sunglasses filter", "UV filter", "SunGlasses RX lenses",
        "Glasses genre", "Glasses usable", "Glasses collection",
        "Items type", "Items packing", "Glasses contain", 
        "Sport glasses", "Glasses frame color effect", "Glasses other features",
        "Glasses clip-on lens color", "Glasses for your face shape", 
        "Glasses lenses no-orders", "Glasses other info"
    ]
    
    # 2. CREATE MISSING COLUMNS
    for col in TARGET_COLUMNS:
        if col not in user_df.columns:
            user_df[col] = "" # Create empty column if missing
            
    # 3. APPLY RULES (We will add your real rules here next!)
    user_df['Items type'] = 'Glasses' # Default Rule
    
    # 4. REORDER
    # Puts the 33 standard columns first, keeps any extra user columns at the end
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
            
            st.success("✅ Done! Data processed.")
            
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