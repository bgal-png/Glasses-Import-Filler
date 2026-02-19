import streamlit as st
import pandas as pd
import os
import re

# ==========================================
# 🛑 VERSION CHECK 
# ==========================================
st.set_page_config(page_title="Manufacturer Data Linker", layout="wide")
APP_VERSION = "3.0 - BARCODE MATCHER & FLIP MAPPING"

st.title(f"🏭 Manufacturer Data Linker")
st.caption(f"🚀 Running Code Version: **{APP_VERSION}**")

# ==========================================
# 🗺️ THE CONFIGURATION (Global -> Manufacturer)
# ==========================================
# Note: Duplicate keys (like multiple lens effects) are now grouped in lists []
MANUFACTURER_CONFIG = {
    "safilo": {
        "file": "safilo.xlsx",
        "brands": ["Carrera", "Polaroid", "Smith", "Boss", "Tommy Hilfiger", "Fossil", "Pierre Cardin", "Marc Jacobs"],
        "columns": {
            "Combination": "Size",
            "Barcode": "EAN/UPC",
            "Glasses_type": "Product Type Desc.",
            "Manufacturer": "Brand Description",
            "Glasses_size_temple_length": "Temple Length",
            "Glasses_size_lens_height": "Lens Height",
            "Glasses_size_lens_width": "Size",
            "Glasses_size_bridge": "Bridge Length",
            "Glasses_shape": "Shape",
            "Glasses_other_info": "Shape",
            "Glasses_frame_type": "Description Rym type",
            "Frame_Colour": "Description Color Family 1",
            "Temple_Colour": "Description Color Family 1",
            "Glasses_main_material": "Group material",
            "Glasses_lens_Colour": "Description Lens Color Family",
            "Glasses_lens_material": "Lens Material Description",
            "Glasses_lens_effect": ["Polarized", "Photochromic", "Treatement Description"],
            "Sunglasses_filter": "Transparency (%)",
            "Glasses_gendre": "Gender",
            "Glasses_model": "Model",
            "Glasses_color_code": "COLOUR CODE",
            "SunGlasses_RX_lenses": "RX able INT",
            "Brand": "Brand Description",
            "Producing_company": ""      
        }
    },
    "luxottica": {
        "file": "luxottica.xlsx",
        "brands": ["Ray-Ban", "Oakley", "Persol", "Prada", "Versace", "Burberry", "Dolce & Gabbana", "Michael Kors", "Vogue", "Arnette", "Ralph"],
        "columns": {
            "Combination": "Velikost",
            "Barcode": "UPC",
            "Glasses_type": "Kolekce",
            "Manufacturer": "Název značky",
            "Glasses_size_temple_length": "Délka stranice",
            "Glasses_size_lens_height": "Výška čočky",
            "Glasses_size_lens_width": "Velikost",
            "Glasses_size_bridge": "Velikost nosníku",
            "Glasses_shape": "Tvar",
            "Glasses_other_info": "Flex",
            "Glasses_frame_type": "Typ",
            "Frame_Colour": "Barva očnice", 
            "Temple_Colour": "Barva očnice", 
            "Glasses_main_material": "Materiál očnice",
            "Glasses_lens_Colour": "Barva čočky",
            "Glasses_lens_material": "Materiál čočky",
            "Glasses_lens_effect": ["Polarizované", "Fotochromatické", "Barva čočky"],
            "Glasses_gendre": "Pohlaví",
            "Glasses_usable": "Motiv",
            "Glasses_collection": "Skládací",
            "Glasses_model": "Kód modelu",
            "Glasses_color_code": "Kód barvy",
            "Brand": "Název značky",
            "Producing_company": "" 
        }
    },
    "kering": {
        "file": "kering.xlsx",
        "brands": ["Gucci", "Saint Laurent", "Balenciaga", "Montblanc", "Bottega Veneta", "Alexander McQueen", "Dunhill", "Puma"],
        "columns": {
            "Combination": "Size 1",
            "Barcode": "UPC code(Z2)",
            "Glasses_type": "Material Group",
            "Manufacturer": "Brand", 
            "Glasses_size_temple_length": "Temple length (mm)",
            "Glasses_size_lens_height": "Lens Height",
            "Glasses_size_lens_width": "Size 1",
            "Glasses_size_bridge": "Bridge Length (mm)",
            "Glasses_shape": "Shape Description",
            "Glasses_other_info": ["Hinge description", "Flex"],
            "Glasses_frame_type": "Rim",
            "Frame_Colour": "Front Main Color Description",
            "Temple_Colour": "Temple Main Color Description",
            "Glasses_main_material": "Front Main Material Description",
            "Glasses_lens_Colour": "Lens Main Color Description",
            "Glasses_lens_material": "Lens Material Description",
            "Glasses_lens_effect": ["Polarized Lens", "Photocromic", "Lens Effect Description"],
            "Sunglasses_filter": "Filter Category",
            "Glasses_gendre": "Fashion Attribute 2",
            "Glasses_collection": "Foldable",
            "SunGlasses_RX_lenses": "Convertible in optical frame",
            "Brand": "Brand",
            "Case_weight_g": "Gross Weight",
            "Glasses_weight_g": "Net Weight",
            "Item_origin_country": "Country of origin",
            "Producing_company": "" 
        }
    },
    "marcolin": {
        "file": "marcolin.xlsx",
        "brands": ["Tom Ford", "Guess", "Adidas", "Max Mara", "Moncler", "Zegna", "Gant", "Harley Davidson", "Skechers", "Web", "Timberland"],
        "columns": {
            "Combination": "Size 1",
            "Barcode": "UPC code(Z2)",
            "Glasses_type": "Material Group",
            "Manufacturer": "Brand", 
            "Glasses_size_temple_length": "Temple length (mm)",
            "Glasses_size_lens_height": "Lens Height",
            "Glasses_size_lens_width": "Size 1",
            "Glasses_size_bridge": "Bridge Length (mm)",
            "Glasses_shape": "Shape Description",
            "Glasses_other_info": ["Hinge description", "Flex"],
            "Glasses_frame_type": "Rim",
            "Frame_Colour": "Front Main Color Description",
            "Temple_Colour": "Temple Main Color Description",
            "Glasses_main_material": "Front Main Material Description",
            "Glasses_lens_Colour": "Lens Main Color Description",
            "Glasses_lens_material": "Lens Material Description",
            "Glasses_lens_effect": ["Polarized Lens", "Photocromic", "Lens Effect Description"],
            "Sunglasses_filter": "Filter Category",
            "Glasses_gendre": "Fashion Attribute 2",
            "Glasses_collection": "Foldable",
            "SunGlasses_RX_lenses": "Convertible in optical frame",
            "Brand": "Brand",
            "Case_weight_g": "Gross Weight",
            "Glasses_weight_g": "Net Weight",
            "Item_origin_country": "Country of origin",
            "Producing_company": "" 
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

        # --- SMART FLIPPER & RENAMER ---
        valid_rename = {}
        cols_to_duplicate = []
        
        for global_name, mfg_names in settings["columns"].items():
            if not mfg_names: continue
            
            # Ensure it's a list for processing
            if isinstance(mfg_names, str):
                mfg_names = [mfg_names]
                
            for mfg_name in mfg_names:
                if mfg_name in df.columns:
                    if mfg_name not in valid_rename:
                        # Standard Rename: MFG_Col -> Global_Col
                        valid_rename[mfg_name] = global_name
                    else:
                        # Handle Multi-Mapping (e.g. Size -> Combination AND Lens_Width)
                        cols_to_duplicate.append((mfg_name, global_name))

        # Perform primary renaming
        df = df.rename(columns=valid_rename)
        
        # Duplicate columns mapped multiple times
        for original_mfg_name, additional_global_name in cols_to_duplicate:
            primary_global_name = valid_rename[original_mfg_name]
            if primary_global_name in df.columns:
                df[additional_global_name] = df[primary_global_name]
                
        # --- THE ULTIMATE BARCODE JOIN KEY ---
        if "Barcode" in df.columns:
            # Strip spaces, make lowercase just in case, and remove trailing '.0' from excel floats
            df["join_key"] = df["Barcode"].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            
            # Optionally remove rows that don't have a barcode
            df = df[df["join_key"].notna() & (df["join_key"] != "nan") & (df["join_key"] != "")]
        else:
            st.error(f"❌ CRITICAL: 'Barcode' column missing in {mfg_name} after rename.")

        # Hardcode the Producer Company directly into the DF
        df["Producing_company"] = mfg_name.title()

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
    # 🕵️ LOGIC CHECK: Do we have Barcode Join Keys?
    first_brand = list(catalog.keys())[0]
    sample_df = catalog[first_brand]
    
    if "join_key" in sample_df.columns:
        st.success(f"✅ SUCCESS: Version {APP_VERSION} active. Barcode Matcher is READY.")
    else:
        st.error(f"❌ FAIL: Barcode 'join_key' is MISSING. Check mappings.")

    st.divider()
    col1, col2 = st.columns([1, 3])
    with col1:
        selected_brand = st.selectbox("Inspect Brand:", sorted(list(catalog.keys())))
    
    with col2:
        if selected_brand:
            df = catalog[selected_brand]
            st.write(f"### Standardized Data for {selected_brand.title()}")
            st.dataframe(df.head(50), use_container_width=True)
            if "join_key" in df.columns:
                st.info(f"Barcode Join Keys Preview: {df['join_key'].head(5).tolist()}")
else:
    st.warning("No catalogs loaded.")