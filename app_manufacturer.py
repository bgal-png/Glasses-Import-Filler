import streamlit as st
import pandas as pd
import os

# 1. Page Configuration
st.set_page_config(page_title="Manufacturer Data Linker", layout="wide")
st.title("🏭 Manufacturer Data Linker: Source Loader")

# ==========================================
# 🗺️ THE CONFIGURATION (Template Mode)
# ==========================================
# I have added "Best Guess" column names here. 
# If they are wrong, the app will error and tell you the REAL names.
MANUFACTURER_CONFIG = {
    "safilo": {
        "file": "safilo.xlsx",
        "brands": ["Carrera", "Polaroid", "Smith", "Boss", "Tommy Hilfiger"], 
        "columns": {
            "Mod.": "model_name",       # Common Safilo header
            "Col.": "color_code",       # Common Safilo header
            "Calibre": "lens_width",    
            "Bridge": "bridge_width"
        }
    },
    "kering": {
        "file": "kering.xlsx",
        "brands": ["Gucci", "Saint Laurent", "Balenciaga", "Montblanc"],
        "columns": {
            "Style": "model_name",      # Common Kering header
            "Color": "color_code",      # Common Kering header
            "Size": "lens_width",
            "Bridge": "bridge_width"
        }
    },
    "marcolin": {
        "file": "marcolin.xlsx",
        "brands": ["Tom Ford", "Guess", "Adidas", "Max Mara"],
        "columns": {
            "Model": "model_name",      # Common Marcolin header
            "Color": "color_code",
            "Eye": "lens_width",
            "Bridge": "bridge_width"
        }
    },
    "luxottica": {
        "file": "luxottica.xlsx",
        "brands": ["Ray-Ban", "Oakley", "Persol", "Prada"],
        "columns": {
            "Model Code": "model_name", # Common Lux header
            "Color Code": "color_code",
            "Size": "lens_width",
            "Bridge": "bridge_width"
        }
    }
}

# ==========================================
# 📥 THE LOADER (Debug Mode)
# ==========================================
@st.cache_data(show_spinner=True)
def load_all_catalogs(config):
    virtual_catalog = {}
    current_dir = os.getcwd()
    
    for mfg_name, settings in config.items():
        file_name = settings["file"]
        file_path = os.path.join(current_dir, file_name)
        
        # 1. Validation
        if not os.path.exists(file_path):
            st.warning(f"⚠️ Missing File: '{file_name}'")
            continue
            
        # 2. Load File (With explicit error printing)
        try:
            if file_name.endswith('.csv'):
                df = pd.read_csv(file_path, dtype=str, on_bad_lines='skip', sep=None, engine='python')
            else:
                # Engine 'openpyxl' is safer for xlsx
                df = pd.read_excel(file_path, dtype=str, engine='openpyxl')
                
            # Normalize headers (strip spaces)
            df.columns = df.columns.astype(str).str.strip()
            
        except Exception as e:
            st.error(f"❌ CRITICAL ERROR loading {file_name}: {e}")
            continue

        # 3. Column Check & Rename
        # We try to rename. If a column is missing, we report it but LOAD THE DATA ANYWAY so you can see it.
        their_cols = list(settings["columns"].keys())
        missing_cols = [c for c in their_cols if c not in df.columns]
        
        if missing_cols:
            st.error(f"❌ {mfg_name.title()} Mapping Error: Columns {missing_cols} not found.")
            st.warning(f"ℹ️ Actual columns in {file_name}: {list(df.columns)[:10]}...") # Show first 10 cols
            # We skip renaming to avoid crash, but we still load the DF so you can inspect it
        else:
            df = df.rename(columns=settings["columns"])
            # Only create join key if rename worked
            if "model_name" in df.columns and "color_code" in df.columns:
                df["join_key"] = df["model_name"].str.strip() + df["color_code"].str.strip()
                df["join_key"] = df["join_key"].str.lower().str.replace(r"[^a-z0-9]", "", regex=True)

        # 4. Success Registration
        for brand in settings["brands"]:
            virtual_catalog[brand.lower().strip()] = df
            
    return virtual_catalog

# ==========================================
# 🚀 APP EXECUTION
# ==========================================
st.sidebar.header("Control Panel")
if st.sidebar.button("🗑️ Clear Cache & Reload"):
    st.cache_data.clear()
    st.rerun()

# Run Loader
if 'catalog' not in st.session_state:
    st.session_state.catalog = load_all_catalogs(MANUFACTURER_CONFIG)

catalog = st.session_state.catalog

# Dashboard
if catalog:
    st.success(f"✅ Loaded data for {len(catalog)} brands.")
    
    st.divider()
    st.subheader("🕵️ Column Inspector")
    
    # Selector
    selected_brand = st.selectbox("Select Brand to Inspect:", sorted(list(catalog.keys())))
    
    if selected_brand:
        df = catalog[selected_brand]
        st.write(f"### Data Preview: {selected_brand.title()}")
        st.dataframe(df.head(50), use_container_width=True)
        
        with st.expander("📋 Copy All Column Names"):
            st.code(list(df.columns))
else:
    st.error("❌ No catalogs loaded. Check the errors above for details.")