import streamlit as st
import pandas as pd
import os
import re
from io import BytesIO

# ==========================================
# 🛑 VERSION CHECK 
# ==========================================
st.set_page_config(page_title="Manufacturer Data Linker", layout="wide")
APP_VERSION = "4.0 - THE MATCHER ACTIVE"

st.title(f"🏭 Manufacturer Data Linker")
st.caption(f"🚀 Running Code Version: **{APP_VERSION}**")

# ==========================================
# 🎯 THE TARGET MAPPING (Your File)
# ==========================================
TARGET_MAPPING = {
    "Combination": "Combination (size on glasses)",
    "Barcode": "Barcode",
    "Glasses_type": "Glasses type ID: 13",
    "Manufacturer": "Manufacturer ID: 9", 
    "Glasses_size_temple_length": "Glasses size: temple length ID: 70",
    "Glasses_size_lens_height": "Glasses size: lens height ID: 71",
    "Glasses_size_lens_width": "Glasses size: lens width ID: 72",
    "Glasses_size_bridge": "Glasses size: bridge ID: 73",
    "Glasses_shape": "Glasses shape ID: 25",
    "Glasses_other_info": "Glasses other info ID: 49",
    "Glasses_frame_type": "Glasses frame type ID: 50",
    "Frame_Colour": "Frame Colour ID: 26",
    "Temple_Colour": "Temple Colour ID: 39",
    "Glasses_main_material": "Glasses main material ID: 53",
    "Glasses_lens_Colour": "Glasses lens Colour ID: 28",
    "Glasses_lens_material": "Glasses lens material ID: 35",
    "Glasses_lens_effect": "Glasses lens effect ID: 37",
    "Sunglasses_filter": "Sunglasses filter ID: 77",
    "Glasses_gendre": "Glasses gendre ID: 22",
    "Glasses_collection": "Glasses collection ID: 33",
    "SunGlasses_RX_lenses": "SunGlasses RX lenses ID:108",
    "Brand": "Brand ID:11",
    "Case_weight_g": "Case weight (g)",
    "Glasses_weight_g": "Glasses weight (g)",
    "Item_origin_country": "Item origin country",
    "Producing_company": "Producing company ID:146" 
}

# ==========================================
# 🗺️ THE CONFIGURATION (Global -> Manufacturer)
# ==========================================
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

        valid_rename = {}
        cols_to_duplicate = []
        
        for global_name, mfg_names in settings["columns"].items():
            if not mfg_names: continue
            if isinstance(mfg_names, str):
                mfg_names = [mfg_names]
                
            for mfg_name_header in mfg_names:
                if mfg_name_header in df.columns:
                    if mfg_name_header not in valid_rename:
                        valid_rename[mfg_name_header] = global_name
                    else:
                        cols_to_duplicate.append((mfg_name_header, global_name))

        df = df.rename(columns=valid_rename)
        
        for original_mfg_name, additional_global_name in cols_to_duplicate:
            primary_global_name = valid_rename[original_mfg_name]
            if primary_global_name in df.columns:
                df[additional_global_name] = df[primary_global_name]
                
        if "Barcode" in df.columns:
            df["join_key"] = df["Barcode"].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            df = df[df["join_key"].notna() & (df["join_key"] != "nan") & (df["join_key"] != "")]
        else:
            st.error(f"❌ CRITICAL: 'Barcode' missing in {mfg_name} after rename.")

        df["Producing_company"] = mfg_name.title()

        for brand in settings["brands"]:
            virtual_catalog[brand.lower().strip()] = df
            
    return virtual_catalog

# ==========================================
# 🚀 APP EXECUTION & UI
# ==========================================

st.sidebar.header("Control Panel")
if st.sidebar.button("🗑️ Clear Cache & Reload Data", type="primary"):
    st.cache_data.clear()
    st.rerun()

# 1. Load Background Data
if 'catalog' not in st.session_state:
    with st.spinner("Building Virtual Catalog..."):
        st.session_state.catalog = load_all_catalogs(MANUFACTURER_CONFIG)

catalog = st.session_state.catalog

if not catalog:
    st.warning("No manufacturer catalogs loaded. Fix errors before proceeding.")
    st.stop()

# 2. Build the Master Database (for fast searching)
@st.cache_data(show_spinner=False)
def get_master_database(cat):
    # Combine all individual brand dataframes into one massive searchable database
    all_dfs = list(cat.values())
    master_df = pd.concat(all_dfs, ignore_index=True)
    # Keep only the first instance of a barcode if it's duplicated across files
    master_df.drop_duplicates(subset=['join_key'], keep='first', inplace=True)
    # Set the barcode as the Index to make searching ultra-fast
    master_df.set_index('join_key', inplace=True)
    return master_df

master_db = get_master_database(catalog)

# 3. Main Interface
st.divider()
st.subheader("📥 Step 1: Upload Your File to Fill")

uploaded_file = st.file_uploader("Upload your Target Excel or CSV file", type=["xlsx", "csv"])

if uploaded_file:
    # Read User File
    try:
        if uploaded_file.name.endswith('.csv'):
            target_df = pd.read_csv(uploaded_file, dtype=str)
        else:
            target_df = pd.read_excel(uploaded_file, dtype=str, engine='openpyxl')
            
        # SANITIZE HEADERS: Remove \n, trailing spaces, etc.
        target_df.columns = target_df.columns.astype(str).str.replace('\n', ' ', regex=False).str.strip()
        
    except Exception as e:
        st.error(f"Could not read your uploaded file: {e}")
        st.stop()

    # Find the target barcode column
    target_barcode_col = TARGET_MAPPING.get("Barcode", "Barcode")
    if target_barcode_col not in target_df.columns:
        st.error(f"❌ Could not find the Barcode column '{target_barcode_col}' in your file. Found columns: {list(target_df.columns)}")
        st.stop()

    st.success(f"File uploaded! Contains {len(target_df)} rows. Click below to start matching.")

    # Execute Matcher
    if st.button("🚀 Run Auto-Filler", type="primary"):
        with st.spinner("Matching barcodes and pouring data..."):
            
            # Ensure target columns exist in the dataframe before we try to write to them
            for global_col, target_col in TARGET_MAPPING.items():
                if target_col not in target_df.columns:
                    target_df[target_col] = "" # Create blank column if missing

            match_count = 0
            
            # Iterate through every row in the user's file
            for index, row in target_df.iterrows():
                raw_barcode = str(row[target_barcode_col]).strip()
                # Clean the barcode exactly like we did in the Master DB
                clean_barcode = re.sub(r'\.0$', '', raw_barcode)
                
                # MAGIC HAPPENS HERE: If we find the barcode in our Master Database
                if clean_barcode in master_db.index:
                    match_count += 1
                    master_row = master_db.loc[clean_barcode]
                    
                    # Pour the data based on mapping
                    for global_col, target_col in TARGET_MAPPING.items():
                        if global_col == "Barcode": continue # Skip overwriting the barcode
                        
                        # Only transfer if the master database has this column and it's not totally empty
                        if global_col in master_db.columns:
                            val = master_row[global_col]
                            if pd.notna(val) and str(val).strip() != "":
                                target_df.at[index, target_col] = val

            st.success(f"✅ Match Complete! Successfully filled {match_count} out of {len(target_df)} products.")
            
            # Preview the result
            st.write("### Preview of Filled Data:")
            st.dataframe(target_df.head(20), use_container_width=True)

            # Generate Downloadable Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                target_df.to_excel(writer, index=False, sheet_name='Filled_Data')
            processed_data = output.getvalue()

            st.download_button(
                label="📥 Download Filled Excel File",
                data=processed_data,
                file_name="Master_Filled_Glasses.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )