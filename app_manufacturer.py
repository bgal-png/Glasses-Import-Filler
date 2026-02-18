import streamlit as st
import pandas as pd
import os
import re

# 1. Page Configuration
st.set_page_config(page_title="Manufacturer Data Linker", layout="wide")
st.title("🏭 Manufacturer Data Linker: Source Loader")
st.write("### 📂 Debug: Files in Current Folder")
import os

st.divider()
st.subheader("🕵️ File Integrity Check")
files_to_check = ["safilo.xlsx", "kering.xlsx", "marcolin.xlsx", "luxottica.xlsx"]

col1, col2 = st.columns(2)
with col1:
    st.write("**File Status:**")
    for f in files_to_check:
        if os.path.exists(f):
            # Get size in Megabytes
            size_mb = os.path.getsize(f) / (1024 * 1024)
            if size_mb < 0.1:
                st.error(f"❌ {f} is too small ({size_mb:.4f} MB). It might be a Git LFS pointer!")
            else:
                st.success(f"✅ {f}: {size_mb:.2f} MB (Looks healthy)")
        else:
            st.warning(f"⚠️ {f} NOT FOUND")

with col2:
    st.write("**Action:**")
    if st.button("🗑️ CLEAR CACHE & RETRY", type="primary"):
        st.cache_data.clear()
        st.rerun()
st.divider()
# ==========================================
# 🗺️ THE CONFIGURATION (Raw Mode)
# ==========================================
# We are loading the files first. You will fill in the "columns" dictionary later.
MANUFACTURER_CONFIG = {
    "safilo": {
        "file": "safilo.xlsx",
        "brands": ["Carrera", "Polaroid", "Smith", "Boss", "Tommy Hilfiger", "Moschino", "Marc Jacobs", "Levi's", "Pierre Cardin"], 
        "columns": {}  # <-- TO BE FILLED LATER
    },
    "kering": {
        "file": "kering.xlsx",
        "brands": ["Gucci", "Saint Laurent", "Balenciaga", "Montblanc", "Bottega Veneta", "Alexander McQueen", "Dunhill"],
        "columns": {}  # <-- TO BE FILLED LATER
    },
    "marcolin": {
        "file": "marcolin.xlsx",
        "brands": ["Tom Ford", "Guess", "Adidas", "Max Mara", "Moncler", "Zegna", "Gant", "Harley Davidson", "Skechers"],
        "columns": {}  # <-- TO BE FILLED LATER
    },
    "luxottica": {
        "file": "luxottica.xlsx",
        "brands": ["Ray-Ban", "Oakley", "Persol", "Prada", "Versace", "Burberry", "Dolce & Gabbana", "Michael Kors", "Vogue"],
        "columns": {}  # <-- TO BE FILLED LATER
    }
}

# ==========================================
# 📥 THE LOADER (Cached for Speed)
# ==========================================
@st.cache_data(show_spinner=True)
def load_all_catalogs(config):
    """
    Loads ALL manufacturer files into a single, standardized 'Virtual Catalog'.
    Returns a Dictionary of DataFrames keyed by 'brand'.
    """
    virtual_catalog = {}
    current_dir = os.getcwd()
    
    # Iterate through every manufacturer defined in the config
    for mfg_name, settings in config.items():
        file_name = settings["file"]
        file_path = os.path.join(current_dir, file_name)
        
        # 1. Check if file exists
        if not os.path.exists(file_path):
            st.warning(f"⚠️ Missing File: '{file_name}' (Skipping {mfg_name})")
            continue
            
        # 2. Load the file (Smart detection of CSV vs Excel)
        try:
            if file_name.endswith('.csv'):
                try:
                    df = pd.read_csv(file_path, dtype=str, on_bad_lines='skip', sep=',')
                except:
                    df = pd.read_csv(file_path, dtype=str, on_bad_lines='skip', sep=';')
            else:
                df = pd.read_excel(file_path, dtype=str)
                
            # Clean up column names (strip whitespace) to avoid "KeyError" later
            df.columns = df.columns.astype(str).str.strip()
            
        except Exception as e:
            st.error(f"❌ Error loading {file_name}: {e}")
            continue

        # 3. Standardize Columns (The Rename) - SKIPPED FOR NOW
        # Once you provide the mappings, we will uncomment this block.
        if settings["columns"]:
            # Check if all target columns exist
            their_cols = list(settings["columns"].keys())
            missing_cols = [c for c in their_cols if c not in df.columns]
            
            if missing_cols:
                st.error(f"❌ Config Error in {mfg_name}: Columns {missing_cols} not found in file.")
                st.write(f"Available columns: {list(df.columns)}")
                continue
                
            df = df.rename(columns=settings["columns"])
        
        # 4. Generate the 'Join Key' (Placeholder)
        # We need "model_name" and "color_code" to do this.
        # Since we haven't mapped them yet, we skip this step safely.
        if "model_name" in df.columns and "color_code" in df.columns:
            df["join_key"] = df["model_name"].str.strip() + df["color_code"].str.strip()
            df["join_key"] = df["join_key"].str.lower().str.replace(r"[^a-z0-9]", "", regex=True)

        # 5. Assign to Brands in the Catalog
        # Map brands to this specific DF
        for brand in settings["brands"]:
            virtual_catalog[brand.lower().strip()] = df
            
    return virtual_catalog

# ==========================================
# 🚀 APP EXECUTION
# ==========================================

st.sidebar.header("Control Panel")
if st.sidebar.button("🔄 Reload Catalogs"):
    st.cache_data.clear()
    st.rerun()

# Load Data
if 'catalog' not in st.session_state:
    with st.spinner("Loading huge manufacturer files..."):
        st.session_state.catalog = load_all_catalogs(MANUFACTURER_CONFIG)

catalog = st.session_state.catalog

# 2. Status Dashboard
if catalog:
    st.success(f"✅ Successfully loaded catalogs for {len(catalog)} brands.")
    
    st.divider()
    st.subheader("🕵️ Data Inspector")
    st.info("Use this section to find the exact Column Headers for mapping.")
    
    col1, col2 = st.columns([1, 3])
    with col1:
        # Create a list of available brands
        available_brands = sorted(list(catalog.keys()))
        selected_brand = st.selectbox("Select a Brand:", available_brands)
    
    with col2:
        if selected_brand:
            df_preview = catalog[selected_brand]
            st.write(f"### Raw Data for {selected_brand.title()} ({len(df_preview)} rows)")
            st.dataframe(df_preview.head(50), use_container_width=True)
            
            # Helper to show all columns easily
            with st.expander("📋 View All Column Names (Copy these for Config)"):
                st.code(list(df_preview.columns))
else:
    st.info("ℹ️ No catalogs loaded yet. Please ensure 'safilo.xlsx', 'kering.xlsx', 'marcolin.xlsx', and 'luxottica.xlsx' are in the folder.")