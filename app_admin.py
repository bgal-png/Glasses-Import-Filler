import streamlit as st
import pandas as pd
from sqlalchemy import create_engine
import os
import re
from dictionaries import (
    MANUFACTURER_CONFIG,
    VALUE_TRANSLATOR,
    FACE_SHAPE_MAP,
    BRAND_USABLE_MAP,
    PREMIUM_KERING_BRANDS,
)

# ==========================================
# 🛑 CONFIG & DATABASE
# ==========================================
st.set_page_config(page_title="Database Admin Panel", layout="wide")
st.title("⚙️ Cloud Database Admin Panel")
st.caption("Upload individual files here to process and merge them into the Supabase Vault.")

# Fetch the secret securely from Streamlit Cloud!
DB_URL = st.secrets["DB_URL"]

@st.cache_resource
def get_engine():
    return create_engine(DB_URL)

@st.cache_resource
def get_engine():
    # pool_pre_ping checks if the connection is alive before sending data!
    return create_engine(DB_URL, pool_pre_ping=True, pool_recycle=300)

engine = get_engine()

# ==========================================
# 🧠 THE ENGINE (ADAPTED FOR UI)
# ==========================================
def load_single_catalog(mfg_name, config_settings, file_path):
    """Runs the massive custom rules engine on a single manufacturer file."""
    unmapped_values = set()
    
    try:
        if file_path.endswith(".csv"):
            try:
                df = pd.read_csv(file_path, dtype=str, on_bad_lines="skip", sep=",")
            except:
                df = pd.read_csv(file_path, dtype=str, on_bad_lines="skip", sep=";")
        else:
            df = pd.read_excel(file_path, dtype=str, engine="openpyxl")

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
        st.error(f"❌ Error loading file: {e}")
        return pd.DataFrame()

    new_df = pd.DataFrame()

    # 1. Map Columns
    for global_name, mfg_names in config_settings["columns"].items():
        if not mfg_names: continue
        if isinstance(mfg_names, str): mfg_names = [mfg_names]
        
        existing_cols = [col for col in mfg_names if col in df.columns]

        if existing_cols:
            if len(existing_cols) == 1:
                col_data = df[existing_cols[0]]
                if isinstance(col_data, pd.DataFrame): col_data = col_data.iloc[:, 0]
                new_df[global_name] = col_data
            else:
                def merge_row(row):
                    vals = [str(row[c]).strip() for c in existing_cols if pd.notna(row[c]) and str(row[c]).strip().lower() not in ("nan", "")]
                    return "|".join(vals) if vals else ""
                new_df[global_name] = df.apply(merge_row, axis=1)

    # 2. Raw Clip-On Engine
    extracted_clip_ons = []
    clip_on_alerts = []

    for idx, raw_row in df.iterrows():
        clip_val = ""
        alert = False

        if mfg_name == "safilo":
            prod_type = re.sub(r"\s+", " ", str(raw_row.get("Product Type Desc.", "")).strip().upper())
            pol = str(raw_row.get("Polarized", "")).strip().upper()
            if "+ CLIP-ON" in prod_type:
                if pol == "0": clip_val = "Magnetic sun clip-on"
                elif pol == "X": clip_val = "Magnetic sun clip-on p"

        elif mfg_name == "luxottica":
            pass  

        elif mfg_name in ["kering", "marcolin"]:
            acc_type = str(raw_row.get("Accessory type", "")).strip().title()
            lens_color = str(raw_row.get("Lens Color mkt", "")).strip().lower()
            sku_desc = str(raw_row.get("SKU Marketing Description", "")).strip().lower()
            pol_lens = str(raw_row.get("Polarized Lens", "")).strip().upper()
            combined_text = lens_color + " " + sku_desc

            if acc_type == "Clip-On":
                if "magnetic clip-on" in combined_text or "magnetic clip on" in combined_text: clip_val = "Magnetic sun clip-on"
                elif "clip-on" in combined_text or "clip on" in combined_text:
                    if "magnetic" not in combined_text: clip_val = "Sun clip-on"
                if clip_val:
                    if "polarized" in combined_text or pol_lens == "X": alert = True

        extracted_clip_ons.append(clip_val)
        clip_on_alerts.append(alert)

    new_df["Extracted_Clip_on"] = extracted_clip_ons
    new_df["Clip_on_Alert"] = clip_on_alerts

    # 3. Custom Rules Strict Engine
    def process_cell_strict(row, col_name, mfg):
        final_values = set()
        raw_val = str(row.get(col_name, "")).strip()

        if col_name == "Glasses_other_info":
            if mfg == "safilo":
                if pd.notna(row.get("Glasses_model")) and "FLEX" in str(row["Glasses_model"]).upper(): final_values.add("Flex")
            elif mfg == "luxottica":
                raw_info = str(row.get("Glasses_other_info", "")).strip().upper()
                if raw_info == "X": final_values.add("Flex")
                if pd.notna(row.get("Glasses_collection")) and str(row["Glasses_collection"]).strip().upper() == "X": final_values.add("Flexible glasses")
            elif mfg in ["kering", "marcolin"]:
                if pd.notna(row.get("Family_descriptions_raw")):
                    if "double bridge" in str(row["Family_descriptions_raw"]).lower(): final_values.add("Double bridge")

        elif col_name == "Glasses_lens_effect":
            if mfg == "safilo":
                if str(row.get("Polarized_raw", "")).strip().upper() == "X": final_values.add("Polarized")
                if str(row.get("Photochromic_raw", "")).strip().upper() == "X": final_values.add("Photochromic")
                raw_eff = str(row.get("Treatement_Description_raw", "")).strip()
                if raw_eff and raw_eff.lower() != "nan":
                    t_dict = VALUE_TRANSLATOR.get(col_name, {})
                    l_dict = {str(k).lower(): v for k, v in t_dict.items() if k}
                    for p in [x.strip() for x in raw_eff.split(",") if x.strip()]:
                        if p.lower() in l_dict:
                            if l_dict[p.lower()]: final_values.add(l_dict[p.lower()])
                        else: unmapped_values.add(f"Safilo -> {col_name}: '{p}'")
            elif mfg == "luxottica":
                if str(row.get("Polarizovane_raw", "")).strip().upper() == "X": final_values.add("Polarized")
                if str(row.get("Fotochromaticke_raw", "")).strip().upper() == "X": final_values.add("Photochromic")
                raw_eff = str(row.get("Barva_cocky_raw", "")).strip()
                if raw_eff and raw_eff.lower() != "nan":
                    matched = False
                    t_dict = VALUE_TRANSLATOR.get(col_name, {})
                    for kw, m_val in t_dict.items():
                        if kw and kw.lower() in raw_eff.lower():
                            if m_val: final_values.add(m_val)
                            matched = True
                    if not matched: unmapped_values.add(f"Luxottica -> {col_name} (Keyword Search): '{raw_eff}'")
            elif mfg in ["kering", "marcolin"]:
                if str(row.get("Polarized_Lens_raw", "")).strip().upper() == "X": final_values.add("Polarized")
                if str(row.get("Photocromic_raw", "")).strip().upper() == "YES": final_values.add("Photochromic")
                raw_eff = str(row.get("Lens_Effect_Description_raw", "")).strip()
                if raw_eff and raw_eff.lower() != "nan":
                    t_dict = VALUE_TRANSLATOR.get(col_name, {})
                    l_dict = {str(k).lower(): v for k, v in t_dict.items() if k}
                    for p in [x.strip() for x in raw_eff.split(",") if x.strip()]:
                        if p.lower() in l_dict:
                            if l_dict[p.lower()]: final_values.add(l_dict[p.lower()])
                        else: unmapped_values.add(f"{mfg.title()} -> {col_name}: '{p}'")
            return "|".join(sorted(list(final_values)))

        elif col_name == "SunGlasses_RX_lenses":
            raw_rx = str(row.get(col_name, "")).strip().upper()
            if mfg in ["safilo", "kering", "marcolin"]:
                if raw_rx == "X": final_values.add("Yes")
            elif mfg == "luxottica": pass
            return "|".join(sorted(list(final_values)))

        elif col_name == "Glasses_shape" and mfg in ["kering", "marcolin"]:
            raw_shape = str(row.get(col_name, "")).strip()
            if raw_shape and raw_shape.lower() != "nan":
                first_shape = raw_shape.split("/")[0].strip()
                if col_name in VALUE_TRANSLATOR:
                    translation_dict = VALUE_TRANSLATOR[col_name]
                    lower_dict = {str(k).lower(): v for k, v in translation_dict.items() if k}
                    shape_lower = first_shape.lower()
                    if shape_lower in lower_dict:
                        if lower_dict[shape_lower]: final_values.add(lower_dict[shape_lower])
                    else: unmapped_values.add(f"{mfg.title()} -> {col_name}: '{first_shape}'")
                else: final_values.add(first_shape)
            return "|".join(sorted(list(final_values)))

        elif col_name == "Sunglasses_filter" and mfg == "safilo":
            raw_val = str(row.get(col_name, "")).strip()
            if raw_val and raw_val.lower() != "nan":
                clean_numbers = re.findall(r"\d+\.?\d*", raw_val)
                matched_by_math = False
                if clean_numbers:
                    vlt = float(clean_numbers[0])
                    if 80 <= vlt <= 100: final_values.add("Category 0"); matched_by_math = True
                    elif 43 <= vlt < 80: final_values.add("Category 1"); matched_by_math = True
                    elif 18 <= vlt < 43: final_values.add("Category 2"); matched_by_math = True
                    elif 8 <= vlt < 18: final_values.add("Category 3"); matched_by_math = True
                    elif 3 <= vlt < 8: final_values.add("Category 4"); matched_by_math = True

                if not matched_by_math:
                    if col_name in VALUE_TRANSLATOR:
                        translation_dict = VALUE_TRANSLATOR[col_name]
                        lower_dict = {str(k).lower(): v for k, v in translation_dict.items() if k}
                        parts = [p.strip() for p in raw_val.split(",") if p.strip()]
                        for part in parts:
                            part_lower = part.lower()
                            if part_lower in lower_dict:
                                if lower_dict[part_lower]: final_values.add(lower_dict[part_lower])
                            else: unmapped_values.add(f"Safilo -> {col_name}: '{part}'")
                    else: final_values.add(raw_val)
            return "|".join(sorted(list(final_values)))

        elif col_name == "Glasses_lens_Colour" and mfg == "luxottica":
            if raw_val and raw_val.lower() != "nan":
                matched = False
                if col_name in VALUE_TRANSLATOR:
                    translation_dict = VALUE_TRANSLATOR[col_name]
                    for keyword, mapped_val in translation_dict.items():
                        if keyword and keyword.lower() in raw_val.lower():
                            if mapped_val: final_values.add(mapped_val)
                            matched = True
                if not matched: unmapped_values.add(f"{mfg.title()} -> {col_name} (Keyword Search): '{raw_val}'")
            return "|".join(sorted(list(final_values)))

        elif raw_val and raw_val.lower() != "nan":
            if col_name in VALUE_TRANSLATOR:
                translation_dict = VALUE_TRANSLATOR[col_name]
                lower_dict = {str(k).lower(): v for k, v in translation_dict.items() if k}
                parts = [p.strip() for p in raw_val.split(",") if p.strip()]
                for part in parts:
                    part_lower = part.lower()
                    if part_lower in lower_dict:
                        if lower_dict[part_lower]: final_values.add(lower_dict[part_lower])
                    else: unmapped_values.add(f"{mfg.title()} -> {col_name}: '{part}'")
            else: final_values.add(raw_val)
        return "|".join(sorted(list(final_values)))

    for target_col in new_df.columns:
        if target_col in VALUE_TRANSLATOR or target_col in ["Glasses_other_info", "Glasses_lens_effect", "SunGlasses_RX_lenses", "Glasses_type", "Glasses_shape", "Sunglasses_filter"]:
            new_df[target_col] = new_df.apply(lambda row: process_cell_strict(row, target_col, mfg_name), axis=1)

    def assemble_name_and_parts(row, mfg):
        brand = str(row.get("Brand", "")).strip().title()
        if brand.lower() == "nan": brand = ""
        model_out, color_out = "", ""

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
                mc_parts = model_color.split(" ")
                model_out = mc_parts[0]
                if len(mc_parts) > 1: color_out = mc_parts[1]
                parts = [brand, model_color]
            else: parts = [brand]
        else: parts = [brand]

        final_name = " ".join([p for p in parts if p])
        return final_name, model_out, color_out

    if not new_df.empty:
        temp_col = new_df.apply(lambda row: assemble_name_and_parts(row, mfg_name), axis=1)
        new_df["Assembled_Name"] = temp_col.apply(lambda x: x[0] if isinstance(x, (list, tuple)) else "")
        new_df["Extracted_Model"] = temp_col.apply(lambda x: x[1] if isinstance(x, (list, tuple)) else "")
        new_df["Extracted_Color"] = temp_col.apply(lambda x: x[2] if isinstance(x, (list, tuple)) else "")

    if "Manufacturer" in new_df.columns:
        new_df["Manufacturer"] = new_df["Manufacturer"].apply(lambda x: str(x).strip().title() if pd.notna(x) and str(x).strip().lower() not in ["nan", ""] else "")
    if "Brand" in new_df.columns:
        new_df["Brand"] = new_df["Brand"].apply(lambda x: str(x).strip().title() if pd.notna(x) and str(x).strip().lower() not in ["nan", ""] else "")

    for dim_col in ["Glasses_size_temple_length", "Glasses_size_lens_height", "Glasses_size_lens_width", "Glasses_size_bridge"]:
        if dim_col in new_df.columns:
            def round_dimension(val):
                if pd.isna(val) or str(val).strip() == "" or str(val).strip().lower() == "nan": return ""
                try:
                    clean_str = re.sub(r"[^\d,.-]", "", str(val).strip()).replace(",", ".")
                    if clean_str: return str(int(round(float(clean_str))))
                except Exception: pass
                return str(val).strip() 
            new_df[dim_col] = new_df[dim_col].apply(round_dimension)

    if "Barcode" in new_df.columns:
        new_df["join_key"] = new_df["Barcode"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True).str.lstrip("0")
        new_df = new_df[new_df["join_key"].notna() & (new_df["join_key"] != "nan") & (new_df["join_key"] != "")]
    else:
        st.error(f"❌ CRITICAL: 'Barcode' missing in {mfg_name} after extraction.")

    new_df["Producing_company"] = mfg_name.title()

    # Expand by brands defined in config
    all_brands_dfs = []
    for brand in config_settings["brands"]:
        brand_df = new_df.copy()
        all_brands_dfs.append(brand_df)

    if unmapped_values:
        with st.expander(f"⚠️ Unmapped Values Found in {mfg_name.title()} File"):
            for val in sorted(unmapped_values): st.write(f"- {val}")

    if all_brands_dfs:
        combined_df = pd.concat(all_brands_dfs, ignore_index=True)
        return combined_df
    return pd.DataFrame()

# ==========================================
# 🔄 UPSERT ENGINE
# ==========================================
def perform_upsert(new_data_df):
    """Takes freshly processed data and intelligently merges it with the Cloud Vault."""
    new_data_df.drop_duplicates(subset=["join_key"], keep="last", inplace=True)
    new_data_df.set_index("join_key", inplace=True)
    
    try:
        existing_df = pd.read_sql_table('master_catalog', con=engine)
        existing_df.set_index("join_key", inplace=True)
        
        # A. Update existing rows with fresh data (if specs changed)
        existing_df.update(new_data_df)
        
        # B. Find brand new rows that don't exist in the database yet
        new_rows = new_data_df[~new_data_df.index.isin(existing_df.index)]
        
        # C. Stitch together
        final_df = pd.concat([existing_df, new_rows])
        upsert_msg = f"✅ Updated existing rows. ✨ Added {len(new_rows)} completely new products!"
    except Exception as e:
        final_df = new_data_df
        upsert_msg = f"✨ Created database from scratch with {len(new_data_df)} products!"

    # Push back to Supabase
    final_df.reset_index().to_sql('master_catalog', con=engine, if_exists='replace', index=False)
    return upsert_msg

# ==========================================
# 🖥️ USER INTERFACE
# ==========================================
st.divider()

col1, col2 = st.columns(2)

# --- 🏭 COLUMN 1: MANUFACTURER CATALOGS ---
with col1:
    st.subheader("🏭 Update Manufacturer Catalogs")
    st.write("Upload a raw manufacturer file. It will be cleaned, processed, and merged into the Master Vault.")
    
    mfg_choice = st.selectbox("Select Manufacturer:", ["safilo", "luxottica", "marcolin", "kering"])
    uploaded_mfg = st.file_uploader(f"Upload new {mfg_choice.title()} file", type=["csv", "xlsx"])
    
    if uploaded_mfg and st.button(f"🚀 Process & Upsert {mfg_choice.title()} Catalog", type="primary"):
        with st.spinner(f"Processing raw {mfg_choice.title()} file through the Rules Engine..."):
            
            # Save uploaded file temporarily to disk so our engine can read it properly
            temp_path = f"temp_{uploaded_mfg.name}"
            with open(temp_path, "wb") as f:
                f.write(uploaded_mfg.getbuffer())
                
            config_settings = MANUFACTURER_CONFIG[mfg_choice]
            processed_df = load_single_catalog(mfg_choice, config_settings, temp_path)
            
            # Clean up temp file
            os.remove(temp_path)
            
            if not processed_df.empty:
                st.success(f"🧹 Cleaned successfully! Extracted {len(processed_df)} raw rows. Pushing to Cloud...")
                msg = perform_upsert(processed_df)
                st.success(msg)
            else:
                st.error("Failed to extract any data. Please check the file format.")

# --- 📦 COLUMN 2: REFERENCE DATA ---
with col2:
    st.subheader("📦 Update Background Data")
    st.write("Upload these files to completely replace their respective tables in the Cloud.")
    
    # Package Data
    pkg_file = st.file_uploader("Upload new `package_data.xlsx`", type=["xlsx"])
    if pkg_file and st.button("⬆️ Replace Package Data in Cloud"):
        with st.spinner("Uploading Package Data..."):
            df_pkg = pd.read_excel(pkg_file)
            df_pkg.columns = df_pkg.columns.astype(str).str.strip()
            df_pkg.to_sql('package_data', engine, if_exists='replace', index=False)
            st.success(f"✅ Package Data updated! ({len(df_pkg)} items)")
            
    st.divider()
    
    # Historical Master Clean
    hist_file = st.file_uploader("Upload new `master_clean.xlsx`", type=["xlsx"])
    if hist_file and st.button("⬆️ Replace Historical Data in Cloud"):
        with st.spinner("Uploading Historical Data..."):
            df_hist = pd.read_excel(hist_file, dtype=str, engine="openpyxl")
            if "Items type" in df_hist.columns:
                df_glasses = df_hist[df_hist["Items type"].astype(str).str.strip().str.lower() == "glasses"]
                df_glasses.columns = df_glasses.columns.astype(str).str.strip()
                df_glasses.to_sql('historical_data', engine, if_exists='replace', index=False)
                st.success(f"✅ Historical Data updated! ({len(df_glasses)} glasses mapped)")
            else:
                st.error("⚠️ 'Items type' column missing. Could not filter/upload historical data.")