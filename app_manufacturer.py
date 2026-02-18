import streamlit as st
import pandas as pd
import os
import re

# ==========================================
# 🛑 VERSION CHECK (The "Spy" Section)
# ==========================================
st.set_page_config(page_title="Manufacturer Data Linker", layout="wide")

# Change this string every time you edit the code to verify updates!
APP_VERSION = "2.1 - DATA MAPPING ACTIVE"

st.title(f"🏭 Manufacturer Data Linker")
st.caption(f"🚀 Running Code Version: **{APP_VERSION}**")

# ==========================================
# 🗺️ THE CONFIGURATION (Exact Mappings)
# ==========================================
MANUFACTURER_CONFIG = {
    "safilo": {
        "file": "safilo.xlsx",
        "brands": ["Carrera", "Polaroid", "Smith", "Boss", "Tommy Hilfiger", "Fossil", "Pierre Cardin", "Marc Jacobs"],
        "columns": {
            "Model": "model_name",
            "COLOUR CODE": "color_code",
            "Size": "lens_width",
            "Bridge Length": "bridge_width",
            "Temple Length": "temple_length",
            "EAN/UPC": "ean",
            "Lens Material Description": "lens_material",
            "Shape": "shape",
            "Gender": "gender"
        }
    },
    "luxottica": {
        "file": "luxottica.xlsx",
        "brands": ["Ray-Ban", "Oakley", "Persol", "Prada", "Versace", "Burberry", "Dolce & Gabbana", "Michael Kors", "Vogue", "Arnette", "Ralph"],
        "columns": {
            "Kód modelu": "model_name",       
            "Kód barvy": "color_code",        
            "Velikost": "lens_width",         
            "Velikost nosníku": "bridge_width",
            "Délka stranice": "temple_length",
            "UPC": "ean",
            "Materiál čočky": "lens_material",
            "Tvar": "shape",
            "Pohlaví": "gender",
            "Materiál stranice": "frame_material"
        }
    },
    "kering": {
        "file": "kering.xlsx",
        "brands": ["Gucci", "Saint Laurent", "Balenciaga", "Montblanc", "Bottega Veneta", "Alexander McQueen", "Dunhill", "Puma"],
        "columns": {
            "SKU Description": "model_raw",   
            "Size 1": "lens_width",
            "Bridge Length (mm)": "bridge_width",
            "Temple length (mm)": "temple_length",
            "UPC code(Z2)": "ean",
            "Lens Material Description": "lens_material",
            "Shape Description": "shape",
            "Fashion Grade": "gender"
        }
    },
    "marcolin": {
        "file": "marcolin.xlsx",
        "brands": ["Tom Ford", "Guess", "Adidas", "Max Mara", "Moncler", "Zegna", "Gant", "Harley Davidson", "Skechers", "Web", "Timberland"],
        "columns": {
            "SKU Description": "model_raw",   
            "Size 1": "lens_width",
            "Bridge Length (mm)": "bridge_width",
            "Temple length (mm)": "temple_length",
            "UPC code(Z2)": "ean",
            "Lens Material Description": "lens_material",
            "Shape Description": "shape"
        }
    }
}

# ==========================================
# 📥 THE LOADER
# ==========================================
@st.cache_data(show_spinner=True)
def load_all_catalogs(config):
    virtual_catalog = {}
    current_dir = os.getcwd()
    
    for mfg_name, settings in config.items():
        file_name = settings["file"]
        file_path = os.path.join(current_dir, file_name)
        
        if not os.path.exists(file_path):
            st.warning(f"⚠️ Missing File: '{file_name}'")
            continue
            
        try:
            if file_name.endswith('.csv'):
                try:
                    df = pd.read_csv(file_path, dtype=str, on_bad_lines='skip', sep=',')
                except:
                    df = pd.read_csv(file_path, dtype=str, on_bad_lines='skip', sep=';')
            else:
                df = pd.read_excel(file_path, dtype=str, engine='openpyxl')
                
            df.columns = df.columns.astype(str).str.strip()
            
        except Exception as e:
            st.error(f"❌ Error loading {file_name}: {e}")
            continue

        # Rename
        valid_rename = {k: v for k, v in settings["columns"].items() if k in df.columns}
        df = df.rename(columns=valid_rename)
        
        # JOIN KEY LOGIC
        if "model_name" in df.columns and "color_code" in df.columns:
            df["join_key"] = df["model_name"].str.strip() + df["color_code"].str.strip()
            
        elif "model_raw" in df.columns:
            def extract_key(raw_val):
                if not isinstance(raw_val, str): return ""
                clean = raw_val.strip().lower()
                return re.sub(r'[^a-z0-9]', '', clean)
            df["join_key"] = df["model_raw"].apply(extract_key)
            
        if "join_key" in df.columns:
            df["join_key"] = df["join_key"].str.lower().str.replace(r"[^a-z0-9]", "", regex=True)

        for brand in settings["brands"]:
            virtual_catalog[brand.lower().strip()] = df
            
    return virtual_catalog

# ==========================================
# 🚀 APP EXECUTION
# ==========================================

st.sidebar.header("Control Panel")
if st.sidebar.button("🗑️ Clear Cache & Reload Data", type="primary"):
    st.cache_data.clear()
    st.rerun()

if 'catalog' not in st.session_state:
    with st.spinner("Building Virtual Catalog..."):
        st.session_state.catalog = load_all_catalogs(MANUFACTURER_CONFIG)

catalog = st.session_state.catalog

if catalog:
    # 🕵️ LOGIC CHECK: Do we have Join Keys?
    first_brand = list(catalog.keys())[0]
    sample_df = catalog[first_brand]
    
    if "join_key" in sample_df.columns:
        st.success(f"✅ SUCCESS: Version {APP_VERSION} is active. 'join_key' column FOUND.")
    else:
        st.error(f"❌ FAIL: Version {APP_VERSION} is active, but 'join_key' is MISSING. Check mappings.")

    st.divider()
    col1, col2 = st.columns([1, 3])
    with col1:
        selected_brand = st.selectbox("Inspect Brand:", sorted(list(catalog.keys())))
    
    with col2:
        if selected_brand:
            df = catalog[selected_brand]
            st.write(f"### Data for {selected_brand.title()}")
            st.dataframe(df.head(50), use_container_width=True)
            if "join_key" in df.columns:
                st.info(f"Join Key Preview: {df['join_key'].head(3).tolist()}")
else:
    st.warning("No catalogs loaded.")