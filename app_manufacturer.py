import streamlit as st
import pandas as pd
import os
import re
from io import BytesIO

# ==========================================
# 🛑 VERSION CHECK 
# ==========================================
st.set_page_config(page_title="Manufacturer Data Linker", layout="wide")
APP_VERSION = "4.2 - ZERO-STRIPPER ACTIVE"

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
# 🔤 THE VALUE TRANSLATOR (Standardizing Data)
# ==========================================
# Format -> "Global_Column_Name": { "Their_Value": "Our_System_Value" }
# IMPORTANT: Put ALL variations from ALL manufacturers in the same list!
VALUE_TRANSLATOR = {
    "Glasses_type": {
        "": "",
    },
    "Manufacturer": {
        #just transfort the text from caps to ccapitalization on the start of each word
    },
    "Glasses_shape": { #na shape až budou kompletní data od výrobců
    #safilo
    "OTHER SHAPE": "Extravagant",
    "ROUND GEOMETRICAL": "Round",
    "ROUND": "Round",
    "PILOT": "Pilot",
    "NAVIGATOR": "Pilot",
    "CAT EYE": "Cat Eye",
    "SQUARE": "Square",
    "SQUARE FLAT TOP": "Square",
    "SQUARE DOUBLE BRIDGE": "Square",
    "SQUARE GEOMETRICAL": "Square",
    "RECTANGULAR GEOMETRICAL": "Rectangular",
    "RECTANGULAR FLAT TOP": "Rectangular",
    "PANTHOS": "Panthos / Tea cup",
    "OVAL": "Oval / Elipse",
    "MASK": "Single lens",
    "BUTTERFLY": "Butterfly",
    "BUTTERFLY GEOMETRICAL": "Butterfly",
    "RECTANGULAR BROWLINE": "Browline",
    #luxottica - napsat dotaz na kategorie
    #kering
    },
    "Glasses_other_info": {
        "": "",
        #safilo
        "SQUARE DOUBLE BRIDGE": "Double bridge",
        #luxottica
        #marcolin
        "Flex":"Flex",
        #kering
        "Flex":"Flex",
    },
    "Glasses_frame_type": {
        "": "",
        #safilo
        #luxottica
        #marcolin
        #kering
    },
    "Frame_Colour": {
        "": "",
        #safilo
        #luxottica
        #marcolin
        #kering
    },
    "Temple_Colour": {
        "": "",
        #safilo
        #luxottica
        #marcolin
        #kering
    },
    "Glasses_main_material": {
        "": "",
        #safilo
        #luxottica
        #marcolin
        #kering
    },
    "Glasses_lens_Colour": {
        "": "",
        #safilo
        #luxottica
        #marcolin
        #kering
    },
    "Glasses_lens_material": {
        "": "",
        #safilo
        #luxottica
        #marcolin
        #kering
    },
    "Glasses_lens_effect": {
        "": "",
        #safilo
        #luxottica
        #marcolin
        #kering
    },
    "Sunglasses_filter": {
        "": "",
        #safilo
        #luxottica
        #marcolin
        #kering
    },
    "Glasses_gendre": {
        "": "",
        #safilo
        #luxottica
        #marcolin
        #kering
    },
    "SunGlasses_RX_lenses": {
        "": "",
        #safilo
        #luxottica
        #marcolin
        #kering
    },
    
    "Glasses_lens_effect": {
        "Dark Grey Mirror Water Polarized": "Polarized, Mirrored" # Example of turning one value into two
    }
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
            "Producing_company": "",
            "Family_descriptions_raw": "Family descriptions",
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
            "Producing_company": "",
            "Family_descriptions_raw": "Family descriptions",
        }
    }
}

# ==========================================
# 📥 THE LOADER
# ==========================================
@st.cache_data(show_spinner=False)
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
            
            new_cols = []
            seen = {}
            for c in df.columns:
                if c in seen:
                    seen[c] += 1
                    new_cols.append(f"{c}.{seen[c]}")
                else:
                    seen[c] = 0
                    new_cols.append(c)
            df.columns = new_cols
            
        except Exception as e:
            st.error(f"❌ Error loading {file_name}: {e}")
            continue

        new_df = pd.DataFrame()
        
        for global_name, mfg_names in settings["columns"].items():
            if not mfg_names: continue
            if isinstance(mfg_names, str):
                mfg_names = [mfg_names]
                
            existing_cols = [col for col in mfg_names if col in df.columns]
            
            if existing_cols:
                if len(existing_cols) == 1:
                    col_data = df[existing_cols[0]]
                    if isinstance(col_data, pd.DataFrame):
                        col_data = col_data.iloc[:, 0]
                    new_df[global_name] = col_data
                else:
                    def merge_row(row):
                        vals = [str(row[c]).strip() for c in existing_cols if pd.notna(row[c]) and str(row[c]).strip().lower() not in ("nan", "")]
                        return ", ".join(vals) if vals else ""
                    new_df[global_name] = df.apply(merge_row, axis=1)

        # ==========================================
        # 🧠 CUSTOM RULES ENGINE & STRICT TRANSLATOR
        # ==========================================
        
        # We will store unknown values here to report to the user later
        if 'unmapped_values' not in st.session_state:
            st.session_state.unmapped_values = set()

        def process_cell_strict(row, col_name, mfg):
            final_values = set()
            
            # --- 1. CUSTOM RULES ENGINE ---
            if col_name == "Glasses_other_info":
                
                # SAFILO RULES
                if mfg == "safilo":
                    if pd.notna(row.get("Glasses_model")) and "FLEX" in str(row["Glasses_model"]).upper():
                        final_values.add("Flex")
                
                # LUXOTTICA RULES
                elif mfg == "luxottica":
                    # Luxottica mapped "Flex" directly to Glasses_other_info initially. 
                    # If it equals "X", we add "Flex"
                    raw_info = str(row.get("Glasses_other_info", "")).strip().upper()
                    if raw_info == "X":
                        final_values.add("Flex")
                    
                    # Check the "Skládací" column (mapped to Glasses_collection)
                    if pd.notna(row.get("Glasses_collection")) and str(row["Glasses_collection"]).strip().upper() == "X":
                        final_values.add("Flexible glasses")
                
                # KERING & MARCOLIN RULES (Identical)
                elif mfg in ["kering", "marcolin"]:
                    if pd.notna(row.get("Family_descriptions_raw")):
                        if "double bridge" in str(row["Family_descriptions_raw"]).lower():
                            final_values.add("Double bridge")

            # --- 2. STRICT DICTIONARY TRANSLATOR ---
            # Now we look at the raw values that came straight from the column mapping
            raw_val = str(row.get(col_name, "")).strip()
            if raw_val and raw_val.lower() != "nan":
                
                # If there is a translator dictionary for this column
                if col_name in VALUE_TRANSLATOR:
                    translation_dict = VALUE_TRANSLATOR[col_name]
                    
                    # Split by comma in case there are multiple values
                    parts = [p.strip() for p in raw_val.split(",") if p.strip()]
                    
                    for part in parts:
                        # Ignore "X" because we handled it in Luxottica custom rules
                        if part.upper() == "X" and col_name == "Glasses_other_info":
                            continue 
                            
                        if part in translation_dict:
                            final_values.add(translation_dict[part])
                        else:
                            # It is NOT specified. Report it and DO NOT fill it.
                            st.session_state.unmapped_values.add(f"{mfg.title()} -> {col_name}: '{part}'")
                else:
                    # If no dictionary exists for this column yet, just keep the raw value
                    final_values.add(raw_val)

            return ", ".join(sorted(list(final_values)))

        # Apply the Engine to all columns we care about
        for target_col in new_df.columns:
            # We only apply this to columns that have rules or dictionaries
            if target_col in VALUE_TRANSLATOR or target_col == "Glasses_other_info":
                new_df[target_col] = new_df.apply(lambda row: process_cell_strict(row, target_col, mfg_name), axis=1)

        # ZERO-STRIPPER APPLIED HERE
        if "Barcode" in new_df.columns:
            new_df["join_key"] = new_df["Barcode"].astype(str).str.strip().str.replace(r'\.0$', '', regex=True).str.lstrip('0')
            new_df = new_df[new_df["join_key"].notna() & (new_df["join_key"] != "nan") & (new_df["join_key"] != "")]
        else:
            st.error(f"❌ CRITICAL: 'Barcode' missing in {mfg_name} after extraction.")

        new_df["Producing_company"] = mfg_name.title()

        for brand in settings["brands"]:
            virtual_catalog[brand.lower().strip()] = new_df
            
    return virtual_catalog

# ==========================================
# 🚀 APP EXECUTION & UI
# ==========================================

st.sidebar.header("Control Panel")

if st.sidebar.button("🗑️ Clear Memory & Reload Data", type="primary"):
    st.cache_data.clear()
    st.session_state.clear() 
    st.rerun()

with st.spinner("Building Virtual Catalog from scratch..."):
    catalog = load_all_catalogs(MANUFACTURER_CONFIG)

if not catalog:
    st.warning("No manufacturer catalogs loaded. Fix errors before proceeding.")
    st.stop()

@st.cache_data(show_spinner=False)
def get_master_database(cat):
    all_dfs = list(cat.values())
    master_df = pd.concat(all_dfs, ignore_index=True)
    master_df.drop_duplicates(subset=['join_key'], keep='first', inplace=True)
    master_df.set_index('join_key', inplace=True)
    return master_df

master_db = get_master_database(catalog)

# 🚨 REPORT UNMAPPED VALUES
if 'unmapped_values' in st.session_state and st.session_state.unmapped_values:
    with st.expander("⚠️ Action Required: Unmapped Values Found!", expanded=True):
        st.warning("The following values were found in the source files but are NOT in your dictionary. They were ignored.")
        for missing in sorted(list(st.session_state.unmapped_values)):
            st.write(f"- {missing}")
        if st.button("Acknowledge & Clear Warnings"):
            st.session_state.unmapped_values = set()
            st.rerun()

st.divider()
st.subheader("📥 Step 1: Upload Your File to Fill")

uploaded_file = st.file_uploader("Upload your Target Excel or CSV file", type=["xlsx", "csv"])

if uploaded_file:
    try:
        if uploaded_file.name.endswith('.csv'):
            target_df = pd.read_csv(uploaded_file, dtype=str)
        else:
            target_df = pd.read_excel(uploaded_file, dtype=str, engine='openpyxl')
            
        target_df.columns = target_df.columns.astype(str).str.replace('\n', ' ', regex=False).str.strip()
        
    except Exception as e:
        st.error(f"Could not read your uploaded file: {e}")
        st.stop()

    target_barcode_col = TARGET_MAPPING.get("Barcode", "Barcode")
    if target_barcode_col not in target_df.columns:
        st.error(f"❌ Could not find the Barcode column '{target_barcode_col}' in your file. Found columns: {list(target_df.columns)}")
        st.stop()

    st.success(f"File uploaded! Contains {len(target_df)} rows. Click below to start matching.")

    if st.button("🚀 Run Auto-Filler", type="primary"):
        with st.spinner("Matching barcodes and pouring data..."):
            
            for global_col, target_col in TARGET_MAPPING.items():
                if target_col not in target_df.columns:
                    target_df[target_col] = "" 

            match_count = 0
            
            for index, row in target_df.iterrows():
                raw_barcode = str(row[target_barcode_col]).strip()
                
                # 🔥 ZERO-STRIPPER APPLIED HERE TOO
                clean_barcode = re.sub(r'\.0$', '', raw_barcode).lstrip('0')
                
                if clean_barcode in master_db.index:
                    match_count += 1
                    master_row = master_db.loc[clean_barcode]
                    
                    for global_col, target_col in TARGET_MAPPING.items():
                        if global_col == "Barcode": continue
                        
                        if global_col in master_db.columns:
                            val = master_row[global_col]
                            if pd.notna(val) and str(val).strip() != "":
                                target_df.at[index, target_col] = val

            st.success(f"✅ Match Complete! Successfully filled {match_count} out of {len(target_df)} products.")
            
            st.write("### Preview of Filled Data:")
            st.dataframe(target_df.head(20), use_container_width=True)

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