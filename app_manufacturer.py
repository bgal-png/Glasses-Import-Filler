import streamlit as st
import pandas as pd
import os
import re
from io import BytesIO
from dictionaries import TARGET_MAPPING, VALUE_TRANSLATOR, MANUFACTURER_CONFIG, FACE_SHAPE_MAP, BRAND_USABLE_MAP, PREMIUM_KERING_BRANDS

# ==========================================
# 🛑 VERSION CHECK 
# ==========================================
st.set_page_config(page_title="Manufacturer Data Linker", layout="wide")
APP_VERSION = "v.260223"

st.title(f"🏭 Manufacturer Data Filler")
st.caption(f"🚀 Running Code Version: **{APP_VERSION}**")

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

                    # --- 📦 PACKAGE DATA UPLOAD ---
        st.markdown("### Step 3: Upload Package Data (Optional)")
        package_file = st.file_uploader("Upload your Package Data file (CSV or Excel)", type=["csv", "xlsx"])
        
        # Initialize an empty dataframe so the app doesn't crash if you don't upload one
        package_df = pd.DataFrame() 
        
        if package_file is not None:
            try:
                # Handle both CSV and Excel formats safely
                if package_file.name.endswith('.csv'):
                    package_df = pd.read_csv(package_file)
                else:
                    package_df = pd.read_excel(package_file)
                    
                st.success(f"✅ Package Data Loaded Successfully! ({len(package_df)} items found)")
                
                with st.expander("👀 Preview Package Data"):
                    st.dataframe(package_df.head())
                    
            except Exception as e:
                st.error(f"❌ Error reading Package file: {e}")

        # ==========================================
        # 🧠 CUSTOM RULES ENGINE & STRICT TRANSLATOR
        # ==========================================
        
        # We will store unknown values here to report to the user later
        if 'unmapped_values' not in st.session_state:
            st.session_state.unmapped_values = set()

        def process_cell_strict(row, col_name, mfg):
            final_values = set()
            raw_val = str(row.get(col_name, "")).strip()
            
            # --- 1. CUSTOM RULES ENGINE ---
            if col_name == "Glasses_other_info":
                if mfg == "safilo":
                    if pd.notna(row.get("Glasses_model")) and "FLEX" in str(row["Glasses_model"]).upper():
                        final_values.add("Flex")
                
                elif mfg == "luxottica":
                    raw_info = str(row.get("Glasses_other_info", "")).strip().upper()
                    if raw_info == "X":
                        final_values.add("Flex")
                    if pd.notna(row.get("Glasses_collection")) and str(row["Glasses_collection"]).strip().upper() == "X":
                        final_values.add("Flexible glasses")
                
                elif mfg in ["kering", "marcolin"]:
                    if pd.notna(row.get("Family_descriptions_raw")):
                        if "double bridge" in str(row["Family_descriptions_raw"]).lower():
                            final_values.add("Double bridge")

            # 🔥 NEW: GLASSES LENS EFFECT ENGINE 🔥
            elif col_name == "Glasses_lens_effect":
                if mfg == "safilo":
                    if str(row.get("Polarized_raw", "")).strip().upper() == "X":
                        final_values.add("Polarized")
                    if str(row.get("Photochromic_raw", "")).strip().upper() == "X":
                        final_values.add("Photochromic")
                    raw_eff = str(row.get("Treatement_Description_raw", "")).strip()
                    if raw_eff and raw_eff.lower() != "nan":
                        t_dict = VALUE_TRANSLATOR.get(col_name, {})
                        l_dict = {str(k).lower(): v for k, v in t_dict.items() if k}
                        for p in [x.strip() for x in raw_eff.split(",") if x.strip()]:
                            if p.lower() in l_dict:
                                if l_dict[p.lower()]: final_values.add(l_dict[p.lower()])
                            else:
                                st.session_state.unmapped_values.add(f"Safilo -> {col_name}: '{p}'")

                elif mfg == "luxottica":
                    if str(row.get("Polarizovane_raw", "")).strip().upper() == "X":
                        final_values.add("Polarized")
                    if str(row.get("Fotochromaticke_raw", "")).strip().upper() == "X":
                        final_values.add("Photochromic")
                    raw_eff = str(row.get("Barva_cocky_raw", "")).strip()
                    if raw_eff and raw_eff.lower() != "nan":
                        matched = False
                        t_dict = VALUE_TRANSLATOR.get(col_name, {})
                        for kw, m_val in t_dict.items():
                            if kw and kw.lower() in raw_eff.lower():
                                if m_val: final_values.add(m_val)
                                matched = True
                        if not matched:
                            st.session_state.unmapped_values.add(f"Luxottica -> {col_name} (Keyword Search): '{raw_eff}'")

                elif mfg in ["kering", "marcolin"]:
                    if str(row.get("Polarized_Lens_raw", "")).strip().upper() == "X":
                        final_values.add("Polarized")
                    if str(row.get("Photocromic_raw", "")).strip().upper() == "YES":
                        final_values.add("Photochromic")
                    raw_eff = str(row.get("Lens_Effect_Description_raw", "")).strip()
                    if raw_eff and raw_eff.lower() != "nan":
                        t_dict = VALUE_TRANSLATOR.get(col_name, {})
                        l_dict = {str(k).lower(): v for k, v in t_dict.items() if k}
                        for p in [x.strip() for x in raw_eff.split(",") if x.strip()]:
                            if p.lower() in l_dict:
                                if l_dict[p.lower()]: final_values.add(l_dict[p.lower()])
                            else:
                                st.session_state.unmapped_values.add(f"{mfg.title()} -> {col_name}: '{p}'")
                
                return ", ".join(sorted(list(final_values)))

            # --- 🕶️ RX LENSES ENGINE ---
            elif col_name == "SunGlasses_RX_lenses":
                raw_rx = str(row.get(col_name, "")).strip().upper()
                
                # Safilo, Kering, and Marcolin all use "X" for Yes
                if mfg in ["safilo", "kering", "marcolin"]:
                    if raw_rx == "X":
                        final_values.add("Yes")
                        
                # Luxottica rule pending (from To-Do list)
                elif mfg == "luxottica":
                    pass 
                
                return ", ".join(sorted(list(final_values)))
            
            # --- 👓 GLASSES SHAPE ENGINE ---
            elif col_name == "Glasses_shape" and mfg in ["kering", "marcolin"]:
                raw_shape = str(row.get(col_name, "")).strip()
                
                if raw_shape and raw_shape.lower() != "nan":
                    # Split the string by '/' and grab only the first item
                    first_shape = raw_shape.split("/")[0].strip()
                    
                    # Run that single extracted shape through the dictionary
                    if col_name in VALUE_TRANSLATOR:
                        translation_dict = VALUE_TRANSLATOR[col_name]
                        lower_dict = {str(k).lower(): v for k, v in translation_dict.items() if k}
                        
                        shape_lower = first_shape.lower()
                        if shape_lower in lower_dict:
                            if lower_dict[shape_lower]: # Add it if it's not mapped to ""
                                final_values.add(lower_dict[shape_lower])
                        else:
                            # Flag only the first shape if it's unknown
                            st.session_state.unmapped_values.add(f"{mfg.title()} -> {col_name}: '{first_shape}'")
                    else:
                        final_values.add(first_shape)
                        
                return ", ".join(sorted(list(final_values)))

           # --- ☀️ SUNGLASSES FILTER ENGINE (Safilo Only) ---
            elif col_name == "Sunglasses_filter" and mfg == "safilo":
                raw_val = str(row.get(col_name, "")).strip()
                
                if raw_val and raw_val.lower() != "nan":
                    clean_numbers = re.findall(r'\d+\.?\d*', raw_val)
                    matched_by_math = False
                    
                    # 1. Try the Math Engine first
                    if clean_numbers:
                        vlt = float(clean_numbers[0])
                        if 80 <= vlt <= 100:
                            final_values.add("Category 0")
                            matched_by_math = True
                        elif 43 <= vlt < 80:
                            final_values.add("Category 1")
                            matched_by_math = True
                        elif 18 <= vlt < 43:
                            final_values.add("Category 2")
                            matched_by_math = True
                        elif 8 <= vlt < 18:
                            final_values.add("Category 3")
                            matched_by_math = True
                        elif 3 <= vlt < 8:
                            final_values.add("Category 4")
                            matched_by_math = True
                            
                    # 2. If math failed (no numbers or out of range), fallback to the Dictionary
                    if not matched_by_math:
                        if col_name in VALUE_TRANSLATOR:
                            translation_dict = VALUE_TRANSLATOR[col_name]
                            lower_dict = {str(k).lower(): v for k, v in translation_dict.items() if k}
                            parts = [p.strip() for p in raw_val.split(",") if p.strip()]
                            
                            for part in parts:
                                part_lower = part.lower()
                                if part_lower in lower_dict:
                                    if lower_dict[part_lower]: # Add if not banned ("")
                                        final_values.add(lower_dict[part_lower])
                                else:
                                    # Still unmapped? Now we flag it!
                                    st.session_state.unmapped_values.add(f"Safilo -> {col_name}: '{part}'")
                        else:
                            final_values.add(raw_val)
                            
                return ", ".join(sorted(list(final_values)))

            # --- 2. KEYWORD SUBSTRING MATCHER (Luxottica Lens Color) ---
            elif col_name == "Glasses_lens_Colour" and mfg == "luxottica":
                if raw_val and raw_val.lower() != "nan":
                    matched = False
                    if col_name in VALUE_TRANSLATOR:
                        translation_dict = VALUE_TRANSLATOR[col_name]
                        for keyword, mapped_val in translation_dict.items():
                            if keyword and keyword.lower() in raw_val.lower():
                                if mapped_val: final_values.add(mapped_val)
                                matched = True
                    if not matched:
                        st.session_state.unmapped_values.add(f"{mfg.title()} -> {col_name} (Keyword Search): '{raw_val}'")
                return ", ".join(sorted(list(final_values)))

            # --- 3. STRICT DICTIONARY TRANSLATOR (Everything else) ---
            elif raw_val and raw_val.lower() != "nan":
                if col_name in VALUE_TRANSLATOR:
                    translation_dict = VALUE_TRANSLATOR[col_name]
                    lower_dict = {str(k).lower(): v for k, v in translation_dict.items() if k}
                    parts = [p.strip() for p in raw_val.split(",") if p.strip()]
                    
                    for part in parts:
                        part_lower = part.lower()
                        if part_lower in lower_dict:
                            if lower_dict[part_lower]: final_values.add(lower_dict[part_lower])
                        else:
                            st.session_state.unmapped_values.add(f"{mfg.title()} -> {col_name}: '{part}'")
                else:
                    final_values.add(raw_val)

            return ", ".join(sorted(list(final_values)))

        # Apply the Engine
        for target_col in new_df.columns:
            if target_col in VALUE_TRANSLATOR or target_col in ["Glasses_other_info", "Glasses_lens_effect", "SunGlasses_RX_lenses", "Glasses_type", "Glasses_shape", "Sunglasses_filter"]:
                new_df[target_col] = new_df.apply(lambda row: process_cell_strict(row, target_col, mfg_name), axis=1)
                # --- 🏷️ NAME ASSEMBLY & EXTRACTION ENGINE ---
        def assemble_name_and_parts(row, mfg):
            brand = str(row.get("Brand", "")).strip().title()
            if brand.lower() == "nan": brand = ""
            
            model_out = ""
            color_out = ""

            if mfg == "safilo":
                model_out = str(row.get("Glasses_model", "")).strip()
                color_out = str(row.get("Glasses_color_code", "")).strip()
                if model_out.lower() == "nan": model_out = ""
                if color_out.lower() == "nan": color_out = ""
                parts = [brand, model_out, color_out]

            elif mfg == "luxottica":
                model_out = str(row.get("Glasses_model", "")).strip().lstrip("0")
                color_out = str(row.get("Glasses_color_code", "")).strip()
                if model_out.lower() == "nan": model_out = ""
                if color_out.lower() == "nan": color_out = ""
                parts = [brand, model_out, color_out]

            elif mfg in ["kering", "marcolin"]:
                mat_num = str(row.get("Material_Number", "")).strip()
                if mat_num and mat_num.lower() != "nan":
                    first_part = mat_num.split(" ")[0]
                    model_color = first_part.replace("-", " ")
                    
                    # Try to split into model and color if there's a space
                    mc_parts = model_color.split(" ")
                    model_out = mc_parts[0]
                    if len(mc_parts) > 1:
                        color_out = mc_parts[1]
                        
                    parts = [brand, model_color]
                else:
                    parts = [brand]
            else:
                parts = [brand]

            final_name = " ".join([p for p in parts if p])
            # 1. Return a simple tuple
            return final_name, model_out, color_out

        # 2. Safely apply and expand the results into the three columns
        # 1. Apply the function and store the result in a temporary column
        if not new_df.empty:
            # We create a temporary column to hold the returned tuple/list
            temp_col = new_df.apply(lambda row: assemble_name_and_parts(row, mfg_name), axis=1)
            
            # 2. Explicitly extract each piece into its own column
            # This bypasses the Pandas 'Columns must be same length as key' check
            new_df["Assembled_Name"] = temp_col.apply(lambda x: x[0] if isinstance(x, (list, tuple)) else "")
            new_df["Extracted_Model"] = temp_col.apply(lambda x: x[1] if isinstance(x, (list, tuple)) else "")
            new_df["Extracted_Color"] = temp_col.apply(lambda x: x[2] if isinstance(x, (list, tuple)) else "")
            
            # Drop the temp column if you want to keep the DF clean
        else:
            new_df["Assembled_Name"] = ""
            new_df["Extracted_Model"] = ""
            new_df["Extracted_Color"] = ""

        # --- 🏭 MANUFACTURER PROPER CASING ---
        if "Manufacturer" in new_df.columns:
            new_df["Manufacturer"] = new_df["Manufacturer"].apply(
                lambda x: str(x).strip().title() if pd.notna(x) and str(x).strip().lower() not in ["nan", ""] else ""
            )
            # --- 🏷️ BRAND PROPER CASING ---
        if "Brand" in new_df.columns:
            new_df["Brand"] = new_df["Brand"].apply(
                lambda x: str(x).strip().title() if pd.notna(x) and str(x).strip().lower() not in ["nan", ""] else ""
            )

            # --- 📏 DIMENSIONS ROUNDING ENGINE ---
        dimension_cols = [
            "Glasses_size_temple_length", 
            "Glasses_size_lens_height", 
            "Glasses_size_lens_width", 
            "Glasses_size_bridge"
        ]
        
        for dim_col in dimension_cols:
            if dim_col in new_df.columns:
                def round_dimension(val):
                    if pd.isna(val) or str(val).strip() == "" or str(val).strip().lower() == "nan":
                        return ""
                    try:
                        # Strip out "mm" or letters, change comma to dot
                        clean_str = re.sub(r'[^\d,.-]', '', str(val).strip()).replace(',', '.')
                        if clean_str:
                            # Convert to float, round to nearest whole number, convert to int string
                            return str(int(round(float(clean_str))))
                    except Exception:
                        pass
                    return str(val).strip() # Fallback to original if math completely fails
                
                new_df[dim_col] = new_df[dim_col].apply(round_dimension)


        # ZERO-STRIPPER
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

st.sidebar.divider()
st.sidebar.subheader("🏷️ Private Name Numbers")
priv_sun = st.sidebar.text_input("Sunglasses", placeholder="e.g. 1001")
priv_eye = st.sidebar.text_input("Eyeglasses (Frames)", placeholder="e.g. 2001")
priv_pc = st.sidebar.text_input("PC Glasses", placeholder="e.g. 3001")
priv_sport = st.sidebar.text_input("Sport Glasses", placeholder="e.g. 4001")
priv_drive = st.sidebar.text_input("Driving Glasses", placeholder="e.g. 5001")

# 🚨 REPORT UNMAPPED VALUES
if 'unmapped_values' in st.session_state and st.session_state.unmapped_values:
    st.warning("⚠️ Action Required: Unmapped Values Found! The following values are not in your dictionary and were ignored.")
    
    # Group the errors by Manufacturer
    unmapped_grouped = {}
    for error in st.session_state.unmapped_values:
        if " -> " in error:
            mfg, detail = error.split(" -> ", 1)
        else:
            mfg, detail = "Other", error
            
        if mfg not in unmapped_grouped:
            unmapped_grouped[mfg] = []
        unmapped_grouped[mfg].append(detail)
        
    # Generate a separate rollout (expander) for each manufacturer
    for mfg in sorted(unmapped_grouped.keys()):
        with st.expander(f"📦 {mfg} Unmapped Values ({len(unmapped_grouped[mfg])})", expanded=False):
            for detail in sorted(unmapped_grouped[mfg]):
                st.write(f"- {detail}")
                
    if st.button("Acknowledge & Clear All Warnings"):
        st.session_state.unmapped_values = set()
        st.rerun()

        # ==========================================
# 🔍 QUICK EAN LOOKUP UTILITY
# ==========================================
st.divider()
st.subheader("🔍 Quick EAN / Barcode Lookup")
col1, col2 = st.columns([3, 1])

with col1:
    search_ean = st.text_input("Enter EAN to search in loaded catalogs:", placeholder="e.g. 8056597123456")
with col2:
    st.write("") # Spacing to align the button
    st.write("")
    search_btn = st.button("Search Database", use_container_width=True)

if search_btn and search_ean:
    # Clean the input exactly like the engine cleans the source barcodes
    clean_search = re.sub(r'\.0$', '', str(search_ean).strip()).lstrip('0')
    
    if clean_search in master_db.index:
        st.success(f"✅ EAN '{search_ean}' found in the database!")
        # Fetch the row(s) and display a few key columns so you know exactly what it is
        found_data = master_db.loc[[clean_search]]
        
        # We will show Manufacturer, Brand, and whatever else is available
        display_cols = [c for c in ["Producing_company", "Brand", "Glasses_type", "Glasses_shape"] if c in found_data.columns]
        st.dataframe(found_data[display_cols], use_container_width=True)
    else:
        st.error(f"❌ EAN '{search_ean}' (Cleaned: {clean_search}) was NOT found in any loaded manufacturer catalog.")

st.divider()
st.subheader("📥 Step 1: Upload Your File to Fill")

uploaded_file = st.file_uploader("Upload your Target Excel or CSV file", type=["xlsx", "csv"])

if uploaded_file:
    try:
        if uploaded_file.name.endswith('.csv'):
            target_df = pd.read_csv(uploaded_file, dtype=str)
        else:
            target_df = pd.read_excel(uploaded_file, dtype=str, engine='openpyxl')
            
        # This strips line breaks, tabs, and squishes multiple spaces into a single space!
        target_df.columns = target_df.columns.astype(str).str.replace('\n', ' ', regex=False).str.replace(r'\s+', ' ', regex=True).str.strip()
        
        # This strips line breaks, tabs, and squishes multiple spaces into a single space!
        target_df.columns = target_df.columns.astype(str).str.replace('\n', ' ', regex=False).str.replace(r'\s+', ' ', regex=True).str.strip()
        
       # 🕵️‍♂️ THE CHEAT CODE: Progress Tracker Edition
        # 1. Grab all the target columns safely (flattens any lists it finds!)
        mapped_targets = set()
        for val in TARGET_MAPPING.values():
            if isinstance(val, list):
                for item in val:
                    mapped_targets.add(item)
            else:
                mapped_targets.add(val)
        
        # 2. List the columns we built custom engines for
        custom_targets = {
            "Items type ID: 20", "Items packing ID: 21", "Name private",
            "Meta description", "Glasses for your face shape ID:94",
            "UV filter ID: 60", "Glasses usable ID: 51", "Glasses collection ID: 33",
            "HS Code", "Item description", "Glasses other features ID:99"
        }
        
        # Combine them into one master list of "Finished" columns
        all_completed_targets = mapped_targets.union(custom_targets)
        
        # 3. Build the visual list
        status_list = []
        for col in target_df.columns:
            if col in all_completed_targets:
                status_list.append(f"✅ {col}")
            else:
                status_list.append(f"⏳ {col}")
                
        # 4. Display in a clean, click-to-open box
        with st.expander("🔍 PROGRESS TRACKER: Exact Bucket Names"):
            st.markdown("**✅ = Rule Applied | ⏳ = Still Needs Logic/Mapping**")
            st.write(status_list)
        
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
            
            # 1. Safely create missing columns (Handles both strings and lists)
            for global_col, target_col in TARGET_MAPPING.items():
                if isinstance(target_col, list):
                    for tc in target_col:
                        if tc not in target_df.columns:
                            target_df[tc] = ""
                else:
                    if target_col not in target_df.columns:
                        target_df[target_col] = "" 

            match_count = 0
            found_sport_glasses = False  # 🚨 Our new tripwire!
            
            for index, row in target_df.iterrows():
                raw_barcode = str(row[target_barcode_col]).strip()
                
                # 🔥 ZERO-STRIPPER
                clean_barcode = re.sub(r'\.0$', '', raw_barcode).lstrip('0')
                
                if clean_barcode in master_db.index:
                    match_count += 1
                    master_row = master_db.loc[clean_barcode]
                    
                    # --- 🛑 FRAMES BYPASS RULE ---
                    # Check if the mapped type is exactly "Frames"
                    is_frames = str(master_row.get("Glasses_type", "")).strip() == "Frames"
                    lens_cols_to_skip = ["Glasses_lens_Colour", "Glasses_lens_material", "Sunglasses_filter", "Glasses_lens_effect"]
                    # --- 📦 STATIC BUCKET FILLS ---
                    target_df.at[index, "Items type ID: 20"] = "Glasses"
                    target_df.at[index, "Items packing ID: 21"] = "Basic"

                    # --- 🕵️ PRIVATE NAME ENGINE ---
                    g_type = str(master_row.get("Glasses_type", "")).strip()
                    private_name = ""
                    
                    # The order here acts as a strict priority hierarchy!
                    if "Sunglasses" in g_type:
                        if priv_sun: private_name = f"(Sunglasses {priv_sun})"
                    elif "Sport glasses" in g_type:
                        if priv_sport: private_name = f"(Sports glasses {priv_sport})"
                    elif "Driving glasses" in g_type:
                        if priv_drive: private_name = f"(Eyeglasses driving {priv_drive})"
                    elif "PC Glasses without power" in g_type:
                        if priv_pc: private_name = f"(Eyeglasses PC {priv_pc})"
                    elif "Frames" in g_type:
                        if priv_eye: private_name = f"(Eyeglasses {priv_eye})"
                        
                    if private_name:
                        target_df.at[index, "Name private"] = private_name.strip()
                    
                    # --- 📝 META DESCRIPTION ENGINE ---
                    assembled_name = str(master_row.get("Assembled_Name", "")).strip()
                    meta_desc = ""
                    
                    if "Sunglasses" in g_type:
                        meta_desc = f"Sunglasses {assembled_name}"
                    elif "Sport glasses" in g_type:
                        meta_desc = f"Ski goggles {assembled_name}"
                        found_sport_glasses = True  # Trip the wire!
                    elif "Driving glasses" in g_type:
                        meta_desc = f"Driving glasses {assembled_name}"
                    elif "PC Glasses without power" in g_type:
                        meta_desc = f"Computer glasses {assembled_name}"
                    elif "Frames" in g_type:
                        meta_desc = f"Eyeglasses {assembled_name}"
                        
                    if meta_desc:
                        target_df.at[index, "Meta description"] = meta_desc.strip()

                    # 2. Safely pour data (Handles both strings and lists)
                    for global_col, target_col in TARGET_MAPPING.items():
                        if global_col == "Barcode": continue
                        
                        # If it's a frame, skip the lens columns entirely
                        if is_frames and global_col in lens_cols_to_skip:
                            continue
                        
                        if global_col in master_db.columns:
                            val = master_row[global_col]
                            if pd.notna(val) and str(val).strip() != "":
                                if isinstance(target_col, list):
                                    for tc in target_col:
                                        target_df.at[index, tc] = val
                                else:
                                    target_df.at[index, target_col] = val
                    # --- 👤 FACE SHAPE ENGINE ---
                    g_shape_raw = str(master_row.get("Glasses_shape", "")).strip()
                    
                    if g_shape_raw and g_shape_raw.lower() not in ["nan", ""]:
                        # Handle potential multiple shapes separated by commas
                        shapes = [s.strip() for s in g_shape_raw.split(",")]
                        recommended_faces = set()
                        
                        for s in shapes:
                            for shape_key, face_val in FACE_SHAPE_MAP.items():
                                if shape_key.lower() == s.lower():
                                    # Split the mapped string by '|' and add to our set to remove duplicates
                                    for face in face_val.split("|"):
                                        recommended_faces.add(face)
                        
                        if recommended_faces:
                            # Join the unique face shapes back together with the pipe separator
                            # (We sort them so they always appear in a clean, alphabetical order!)
                            target_df.at[index, "Glasses for your face shape ID:94"] = "|".join(sorted(list(recommended_faces)))
                    
                    # --- ☀️ UV FILTER ENGINE ---
                    if "Sunglasses" in g_type:
                        target_df.at[index, "UV filter ID: 60"] = "400"
                    
                    # --- 🎯 GLASSES USABLE ENGINE ---
                    usable_tags = set()
                    
                    # 1. Brand Logic
                    raw_brand = str(master_row.get("Brand", "")).strip().lower()
                    if raw_brand in BRAND_USABLE_MAP:
                        usable_tags.add(BRAND_USABLE_MAP[raw_brand])
                        
                    # 2. Polarized / Sunglasses Logic
                    lens_effect = str(master_row.get("Glasses_lens_effect", "")).strip()
                    
                    if "Sunglasses" in g_type:
                        if "Polarized" in lens_effect:
                            usable_tags.add("Driving glasses")
                        else:
                            usable_tags.add("Common use")
                            
                    # 3. Join and Pour
                    if usable_tags:
                        # Sorting them ensures it always looks clean, e.g., "Common use|Luxury glasses"
                        target_df.at[index, "Glasses usable ID: 51"] = "|".join(sorted(list(usable_tags)))

                    # --- 💎 PREMIUM COLLECTION ENGINE ---
                    # We reuse the 'raw_brand' variable from the Usable Engine above!
                    if raw_brand in PREMIUM_KERING_BRANDS:
                        target_df.at[index, "Glasses collection ID: 33"] = "Prémiové brýle - Kering"

                        # --- 🌍 HS CODE ENGINE ---
                    # We reuse 'g_type' from the very beginning of the loop!
                    raw_material = str(master_row.get("Glasses_main_material", "")).strip().lower()
                    
                    if "Sunglasses" in g_type:
                        target_df.at[index, "HS Code"] = "90041091"
                    elif "Sport glasses" in g_type:
                        target_df.at[index, "HS Code"] = "90049090"
                    elif "Frames" in g_type:
                        if "plastic" in raw_material:
                            target_df.at[index, "HS Code"] = "90031100"
                        elif "metal" in raw_material:
                            target_df.at[index, "HS Code"] = "90031900"
                    
                    # --- 📝 ITEM DESCRIPTION ENGINE ---
                    # We continue to reuse 'g_type' and 'raw_material' from above!
                    if "Frames" in g_type:
                        target_df.at[index, "Item description"] = "Eyeglasses"
                    elif "PC Glasses without power" in g_type:
                        target_df.at[index, "Item description"] = "PC Glasses without power"
                    elif "Driving glasses" in g_type:
                        target_df.at[index, "Item description"] = "Driving glasses"
                    elif "Sunglasses" in g_type:
                        has_plastic = "plastic" in raw_material
                        has_metal = "metal" in raw_material
                        
                        if has_plastic and has_metal:
                            target_df.at[index, "Item description"] = "Sunglasses, mixed plastic and metal frame"
                        elif has_plastic:
                            target_df.at[index, "Item description"] = "Sunglasses, plastic frame"
                        elif has_metal:
                            target_df.at[index, "Item description"] = "Sunglasses, metal frame"
                            # ---------------------------------------------------------
                    # THE FOLLOWING RULES MUST RUN *AFTER* THE GENERIC POUR
                    # ---------------------------------------------------------

                    # --- 🌟 OTHER FEATURES ENGINE ---
                    other_features = set()
                    
                    # If the column already received data from the generic pour, grab it so we don't overwrite it!
                    if "Glasses other features ID:99" in target_df.columns:
                        existing_features = str(target_df.at[index, "Glasses other features ID:99"]).strip()
                        if existing_features and existing_features.lower() not in ["nan", ""]:
                            for e in existing_features.split("|"):
                                other_features.add(e.strip())
                    
                    # 1. Check RX Lenses
                    if "SunGlasses RX lenses ID:108" in target_df.columns:
                        rx_val = str(target_df.at[index, "SunGlasses RX lenses ID:108"]).strip().lower()
                        if rx_val == "yes":
                            other_features.add("Prescription sunglasses")
                            
                    # 2. Check Clip-ons
                    if "Glasses contain ID: 84" in target_df.columns:
                        contain_val = str(target_df.at[index, "Glasses contain ID: 84"]).strip().lower()
                        
                        # Split by comma or pipe to isolate the exact phrases 
                        # (Prevents "magnetic sun clip-on" from accidentally triggering the basic "sun clip-on" rule!)
                        contain_items = [item.strip() for item in re.split(r'[,|]', contain_val) if item.strip()]
                        
                        clip_on_found = False
                        if "sun clip-on" in contain_items:
                            other_features.add("Sun clip-on")
                            clip_on_found = True
                        if "sun clip-on p" in contain_items:
                            other_features.add("Sun clip-on p")
                            clip_on_found = True
                        if "magnetic sun clip-on" in contain_items:
                            other_features.add("Magnetic sun clip-on")
                            clip_on_found = True
                        if "magnetic sun clip-on p" in contain_items:
                            other_features.add("Magnetic sun clip-on p")
                            clip_on_found = True
                            
                        # 3. Add the universal clip-on tag
                        if clip_on_found:
                            other_features.add("Glasses with sun clip-on")
                            
                    # 4. Pour into target
                    if other_features:
                        target_df.at[index, "Glasses other features ID:99"] = "|".join(sorted(list(other_features)))



            st.success(f"✅ Match Complete! Successfully filled {match_count} out of {len(target_df)} products.")

            # 🚨 Trigger the Sport Glasses Warning if the tripwire was crossed
            if found_sport_glasses:
                st.warning("⚠️ **Heads Up:** We found 'Sport glasses' in this batch and labeled them as 'Ski goggles' in the Meta Description. Please double-check the final file to ensure they aren't cycling or swimming glasses!")
            
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