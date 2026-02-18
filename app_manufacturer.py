import streamlit as st
import pandas as pd
import os
import re

# 1. Page Configuration
st.set_page_config(page_title="Manufacturer Data Linker", layout="wide")
st.title("🏭 Manufacturer Data Linker: The 'Source of Truth'")

# ==========================================
# 🗺️ THE CONFIGURATION (The Rosetta Stone)
# ==========================================
# UPDATE THIS with your real filenames and column headers!
MANUFACTURER_CONFIG = {
    "safilo": {
        "file": "safilo_catalog.xlsx",  # <--- Change to your real filename
        "brands": ["Carrera", "Polaroid", "Smith", "Boss", "Tommy Hilfiger"],
        "columns": {
            "Mod.": "model_name",       # Their Column -> Our Standard Key
            "Col.": "color_code",
            "Calibre": "lens_width",
            "Bridge": "bridge_width",
            "Temple": "temple_length"
        }
    },
    "kering": {
        "file": "kering_master.csv",    # <--- Change to your real filename
        "brands": ["Gucci", "Saint Laurent", "Balenciaga", "Montblanc"],
        "columns": {
            "Style": "model_name",
            "ColorId": "color_code",
            "Size": "lens_width",
            "Bridge": "bridge_width",
            "TempleLength": "temple_length"
        }
    },
    # Add other manufacturers here following the same pattern
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
                # Try reading with different delimiters just in case
                try:
                    df = pd.read_csv(file_path, dtype=str, on_bad_lines='skip', sep=',')
                except:
                    df = pd.read_csv(file_path, dtype=str, on_bad_lines='skip', sep=';')
            else:
                df = pd.read_excel(file_path, dtype=str)
        except Exception as e:
            st.error(f"❌ Error loading {file_name}: {e}")
            continue

        # 3. Standardize Columns (The Rename)
        # Check if all columns exist before renaming
        their_cols = list(settings["columns"].keys())
        missing_cols = [c for c in their_cols if c not in df.columns]
        
        if missing_cols:
            st.error(f"❌ Config Error in {mfg_name}: Columns {missing_cols} not found in file. Check your config!")
            st.write(f"Available columns in {file_name}: {list(df.columns)}")
            continue
            
        # Rename to our standard names
        df = df.rename(columns=settings["columns"])
        
        # 4. Generate the 'Join Key' (Critical Step!)
        # We create a standardized fingerprint: Brand + Model + Color
        # Logic: Lowercase, remove spaces, remove dashes/slashes
        if "model_name" in df.columns and "color_code" in df.columns:
            df["join_key"] = df["model_name"].str.strip() + df["color_code"].str.strip()
            df["join_key"] = df["join_key"].str.lower().str.replace(r"[^a-z0-9]", "", regex=True)

        # 5. Assign to Brands in the Catalog
        # Instead of storing one huge DF, we map brands to this specific DF
        # This makes lookup instantaneous: virtual_catalog["Gucci"] -> Kering DF
        for brand in settings["brands"]:
            virtual_catalog[brand.lower().strip()] = df
            
    return virtual_catalog

# ==========================================
# 🚀 APP EXECUTION
# ==========================================

st.write("### 1. Catalog Status")

if st.button("🔄 Reload Catalogs"):
    st.cache_data.clear()
    st.rerun()

# Load Data
if 'catalog' not in st.session_state:
    st.session_state.catalog = load_all_catalogs(MANUFACTURER_CONFIG)

catalog = st.session_state.catalog

# 2. Status Dashboard
if catalog:
    st.success(f"✅ Successfully loaded catalogs for {len(catalog)} brands.")
    
    col1, col2 = st.columns([1, 3])
    with col1:
        selected_brand = st.selectbox("Select a Brand to Inspect Source Data:", list(catalog.keys()))
    
    with col2:
        if selected_brand:
            st.write(f"### Source Data for {selected_brand.title()}")
            st.dataframe(catalog[selected_brand].head(50), use_container_width=True)
else:
    st.info("ℹ️ No catalogs loaded yet. Please add files to the folder and update MANUFACTURER_CONFIG.")