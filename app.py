import streamlit as st
import pandas as pd
import os
import io

# 1. Page Config
st.set_page_config(page_title="Excel Auto-Filler", layout="wide")
st.title("⚡ Excel Data Filler: Glasses Edition")

# ==========================================
# 🔒 LOADER: Uses 'master_clean.xlsx'
# (Same Indestructible Logic as Validator)
# ==========================================
@st.cache_data
def load_master():
    """Reads the master database to use for lookups."""
    current_dir = os.getcwd()
    candidates = [f for f in os.listdir(current_dir) if (f.endswith('.xlsx') or f.endswith('.csv')) and "master_clean" in f and not f.startswith('~$')]
    
    if not candidates:
        st.error("❌ 'master_clean.xlsx' not found in repository."); st.stop()
    
    file_path = candidates[0]
    df = None
    
    try:
        df = pd.read_excel(file_path, dtype=str, engine='openpyxl')
    except Exception:
        # Fallback to CSV strategies
        strategies = [{'sep': None}, {'sep': ','}, {'sep': ';'}]
        for strat in strategies:
            try:
                df = pd.read_csv(file_path, dtype=str, on_bad_lines='skip', **strat)
                break
            except: continue
            
    if df is None:
        st.error("❌ Could not read Master File."); st.stop()
        
    # Standardize Headers
    df.columns = df.columns.astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
    return df

# Load Master immediately
master_df = load_master()
st.success(f"✅ Brain Loaded: {len(master_df)} rows from Master Database")

# ==========================================
# 🧠 THE BRAIN: FILLING LOGIC
# This is where we will write your specific rules later
# ==========================================
def run_auto_fill(user_df, master_df):
    """
    Takes the user's partial data and fills in the blanks
    based on the Master Data rules.
    """
    # 1. DEFINE THE TARGET STRUCTURE (The 33 Columns)
    # These are the columns the Validator expects.
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
    # If the user uploaded a file with just "Brand" and "Model",
    # we add the other 31 columns as empty placeholders first.
    for col in TARGET_COLUMNS:
        if col not in user_df.columns:
            user_df[col] = "" # Create empty column
            
    # 3. APPLY RULES (EXAMPLES - WE WILL CHANGE THESE)
    # ------------------------------------------------
    # Rule Example 1: Always set 'Items type' to 'Glasses'
    user_df['Items type'] = 'Glasses'
    
    # Rule Example 2: If 'Brand' is filled, try to find 'Manufacturer'
    # (This is just a placeholder logic to show it working)
    if 'Brand' in user_df.columns:
         # Simple lookup logic could go here
         pass
         
    # ------------------------------------------------
    
    # 4. REORDER COLUMNS
    # Ensure the final file is in the nice order we defined above
    # (Any extra columns the user added, like 'Internal ID', stay at the end)
    final_cols = TARGET_COLUMNS + [c for c in user_df.columns if c not in TARGET_COLUMNS]
    return user_df[final_cols]

# ==========================================
# 📤 USER INTERFACE
# ==========================================
st.divider()
st.subheader("1. Upload Partial Data")
st.info("Upload your Excel file. I will ensure all 33 columns exist and fill what I can.")

uploaded_file = st.file_uploader("Choose Excel File", type=['xlsx'])

if uploaded_file:
    # Load User Data
    try:
        user_df = pd.read_excel(uploaded_file, dtype=str)
    except:
        user_df = pd.read_csv(uploaded_file, dtype=str)
        
    st.write(f"Loaded {len(user_df)} rows. Columns found: {list(user_df.columns)}")

    st.divider()
    st.subheader("2. Run Auto-Fill")
    
    if st.button("✨ Auto-Fill Data", type="primary"):
        with st.spinner("Applying rules..."):
            # RUN THE BRAIN
            filled_df = run_auto_fill(user_df, master_df)
            
            st.success("✅ Done! Data has been standardized.")
            
            # PREVIEW
            st.dataframe(filled_df.head())
            
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