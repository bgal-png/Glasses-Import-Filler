import streamlit as st
import pandas as pd
import os
import re

# 1. Page Configuration
st.set_page_config(page_title="Manufacturer Data Linker", layout="wide")
st.title("🏭 Manufacturer Data Linker: Source Loader")

# ==========================================
# 🗺️ THE CONFIGURATION (Exact Mappings from your Files)
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
            "Kód modelu": "model_name",       # Czech Header
            "Kód barvy": "color_code",        # Czech Header
            "Velikost": "lens_width",         # Czech Header
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
            "SKU Description": "model_raw",   # Contains "Model-Color" (e.g., AM0001S-001)
            "Size 1": "lens_width",
            "Bridge Length (mm)": "bridge_width",
            "Temple length (mm)": "temple_length",
            "UPC code(Z2)": "ean",
            "Lens Material Description": "lens_material",
            "Shape Description": "shape",
            "Fashion Grade": "gender"         # Sometimes gender is hidden here or in 'Concept'
        }
    },
    "marcolin": {
        "file": "marcolin.xlsx",
        "brands": ["Tom Ford", "Guess", "Adidas", "Max Mara", "Moncler", "Zegna", "Gant", "Harley Davidson", "Skechers", "Web", "Timberland"],
        "columns": {
            "SKU Description": "model_raw",   # Contains "Model-Color"
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
            # Load File (Smart Engine Selection)
            if file_name.endswith('.csv'):
                # Try standard comma first, then semicolon
                try:
                    df = pd.read_csv(file_path, dtype=str, on_bad_lines='skip', sep=',')
                except:
                    df = pd.read_csv(file_path, dtype=str, on_bad_lines='skip', sep=';')
            else:
                df = pd.read_excel(file_path, dtype=str, engine='openpyxl')
                
            # Normalize Headers (Strip spaces)
            df.columns = df.columns.astype(str).str.strip()
            
        except Exception as e:
            st.error(f"❌ Error loading {file_name}: {e}")
            continue

        # Rename Columns
        # Check if map exists to avoid KeyErrors
        available_cols = set(df.columns)
        required_cols = set(settings["columns"].keys())
        valid_rename = {k: v for k, v in settings["columns"].items() if k in available_cols}
        
        # Report Missing Columns (Debug Info)
        missing = required_cols - available_cols
        if missing:
            st.warning(f"⚠️ {mfg_name.title()}: Could not find columns {missing}. Check file format.")
        
        df = df.rename(columns=valid_rename)
        
        # --- JOIN KEY GENERATION LOGIC ---
        
        # Logic A: Standard (Safilo/Luxottica) - Separate Columns
        if "model_name" in df.columns and "color_code" in df.columns:
            df["join_key"] = df["model_name"].str.strip() + df["color_code"].str.strip()
            
        # Logic B: Kering/Marcolin - Combined in "model_raw" (e.g., "GG0001S-001")
        elif "model_raw" in df.columns:
            # We assume format is "MODEL-COLOR" or similar.
            # We split by space or hyphen to get components.
            # This is a basic extractor; we can refine it later.
            def extract_key(raw_val):
                if not isinstance(raw_val, str): return ""
                # Remove spaces and normalize
                clean = raw_val.strip().lower()
                # Remove non-alphanumeric chars to make a "mush" key
                return re.sub(r'[^a-z0-9]', '', clean)
            
            df["join_key"] = df["model_raw"].apply(extract_key)
            
        # Clean Key (Lowercase, no special chars)
        if "join_key" in df.columns:
            df["join_key"] = df["join_key"].str.lower().str.replace(r"[^a-z0-9]", "", regex=True)

        # Register to Virtual Catalog
        for brand in settings["brands"]:
            virtual_catalog[brand.lower().strip()] = df
            
    return virtual_catalog

# ==========================================
# 🚀 APP EXECUTIONS
# ==========================================

st.sidebar.header("Control Panel")
if st.sidebar.button("🔄 Reload Catalogs"):
    st.cache_data.clear()
    st.rerun()

if 'catalog' not in st.session_state:
    with st.spinner("Creating Virtual Catalog..."):
        st.session_state.catalog = load_all_catalogs(MANUFACTURER_CONFIG)

catalog = st.session_state.catalog

if catalog:
    st.success(f"✅ System Online: Loaded {len(catalog)} brands.")
    
    col1, col2 = st.columns([1, 3])
    with col1:
        selected_brand = st.selectbox("Inspect Brand:", sorted(list(catalog.keys())))
    
    with col2:
        if selected_brand:
            df = catalog[selected_brand]
            st.write(f"### Standardized Data: {selected_brand.title()}")
            st.dataframe(df.head(50), use_container_width=True)
            
            st.info(f"Join Keys generated: {df['join_key'].head(3).values if 'join_key' in df.columns else 'None'}")