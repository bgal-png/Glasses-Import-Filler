import streamlit as st
import pandas as pd
import os
import re
from io import BytesIO
from dictionaries import TARGET_MAPPING, VALUE_TRANSLATOR, MANUFACTURER_CONFIG

# ==========================================
# 🛑 VERSION CHECK 
# ==========================================
st.set_page_config(page_title="Manufacturer Data Linker", layout="wide")
APP_VERSION = "4.2 - ZERO-STRIPPER ACTIVE"

st.title(f"🏭 Manufacturer Data Linker")
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
            if target_col in VALUE_TRANSLATOR or target_col in ["Glasses_other_info", "Glasses_lens_effect", "SunGlasses_RX_lenses"]:
                new_df[target_col] = new_df.apply(lambda row: process_cell_strict(row, target_col, mfg_name), axis=1)
                # Apply the Engine
        for target_col in new_df.columns:
            if target_col in VALUE_TRANSLATOR or target_col in ["Glasses_other_info", "Glasses_lens_effect", "SunGlasses_RX_lenses", "Glasses_shape"]:
                new_df[target_col] = new_df.apply(lambda row: process_cell_strict(row, target_col, mfg_name), axis=1)

        # --- 🏷️ NAME ASSEMBLY ENGINE ---
        def assemble_name(row, mfg):
            # 1. Get and format the Brand (Title Case / Proper)
            brand = str(row.get("Brand", "")).strip().title()
            if brand.lower() == "nan": brand = ""

            # 2. Manufacturer Specific Logic
            if mfg == "safilo":
                model = str(row.get("Glasses_model", "")).strip()
                color = str(row.get("Glasses_color_code", "")).strip()
                if model.lower() == "nan": model = ""
                if color.lower() == "nan": color = ""
                
                parts = [brand, model, color]

            elif mfg == "luxottica":
                model = str(row.get("Glasses_model", "")).strip()
                # Remove leading zeros from model
                model = model.lstrip("0")
                if model.lower() == "nan": model = ""
                
                color = str(row.get("Glasses_color_code", "")).strip()
                if color.lower() == "nan": color = ""
                
                parts = [brand, model, color]

            elif mfg in ["kering", "marcolin"]:
                mat_num = str(row.get("Material_Number", "")).strip()
                if mat_num and mat_num.lower() != "nan":
                    # Get everything before the first space
                    first_part = mat_num.split(" ")[0]
                    # Replace hyphen with space
                    model_color = first_part.replace("-", " ")
                    parts = [brand, model_color]
                else:
                    parts = [brand]
            else:
                parts = [brand]

            # Join parts with a space, filtering out any empty strings
            return " ".join([p for p in parts if p])

        # Apply the name builder to create the new column
        new_df["Assembled_Name"] = new_df.apply(lambda row: assemble_name(row, mfg_name), axis=1)


        # ZERO-STRIPPER
        if "Barcode" in new_df.columns:

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