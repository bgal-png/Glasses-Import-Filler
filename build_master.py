import pandas as pd
from sqlalchemy import create_engine
import re
from dictionaries import TARGET_MAPPING, VALUE_TRANSLATOR, MANUFACTURER_CONFIG, FACE_SHAPE_MAP, BRAND_USABLE_MAP, PREMIUM_KERING_BRANDS

# ⚠️ REMEMBER: Change your password in Supabase after we are done!
DB_URL = "postgresql://postgres.nxlwkzgfcmzsbogcenyi:YQe2oULo6y6WXOZN@aws-1-eu-central-1.pooler.supabase.com:6543/postgres"

def build_master_catalog():
    print("🚀 Connecting to Supabase Vault to build Master Catalog...")
    engine = create_engine(DB_URL)
    virtual_catalog = {}
    unmapped_values = set()

    for mfg_name, settings in MANUFACTURER_CONFIG.items():
        table_name = f"raw_{mfg_name.lower()}"
        print(f"📥 Pulling raw data from '{table_name}'...")
        
        try:
            # 1. READ RAW DATA DIRECTLY FROM THE VAULT
            df = pd.read_sql_table(table_name, con=engine)
        except Exception as e:
            print(f"⚠️ Skipping {mfg_name.title()}: Table '{table_name}' not found in database.")
            continue

        new_df = pd.DataFrame()
        
        # 2. MAP COLUMNS
        for global_name, mfg_names in settings["columns"].items():
            if not mfg_names: continue
            if isinstance(mfg_names, str):
                mfg_names = [mfg_names]
                
            existing_cols = [col for col in mfg_names if col in df.columns]
            
            if existing_cols:
                if len(existing_cols) == 1:
                    new_df[global_name] = df[existing_cols[0]]
                else:
                    def merge_row(row):
                        vals = [str(row[c]).strip() for c in existing_cols if pd.notna(row[c]) and str(row[c]).strip().lower() not in ("nan", "")]
                        return ", ".join(vals) if vals else ""
                    new_df[global_name] = df.apply(merge_row, axis=1)

        # 3. APPLY CUSTOM RULES ENGINE
        def process_cell_strict(row, col_name, mfg):
            final_values = set()
            raw_val = str(row.get(col_name, "")).strip()
            
            if col_name == "Glasses_other_info":
                if mfg == "safilo":
                    if pd.notna(row.get("Glasses_model")) and "FLEX" in str(row["Glasses_model"]).upper():
                        final_values.add("Flex")
                elif mfg == "luxottica":
                    raw_info = str(row.get("Glasses_other_info", "")).strip().upper()
                    if raw_info == "X": final_values.add("Flex")
                    if pd.notna(row.get("Glasses_collection")) and str(row["Glasses_collection"]).strip().upper() == "X":
                        final_values.add("Flexible glasses")
                elif mfg in ["kering", "marcolin"]:
                    if pd.notna(row.get("Family_descriptions_raw")):
                        if "double bridge" in str(row["Family_descriptions_raw"]).lower():
                            final_values.add("Double bridge")

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
                        if not matched: unmapped_values.add(f"Luxottica -> {col_name}: '{raw_eff}'")
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
                return ", ".join(sorted(list(final_values)))

            elif col_name == "SunGlasses_RX_lenses":
                raw_rx = str(row.get(col_name, "")).strip().upper()
                if mfg in ["safilo", "kering", "marcolin"] and raw_rx == "X":
                    final_values.add("Yes")
                return ", ".join(sorted(list(final_values)))
            
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
                return ", ".join(sorted(list(final_values)))

            elif col_name == "Sunglasses_filter" and mfg == "safilo":
                raw_val = str(row.get(col_name, "")).strip()
                if raw_val and raw_val.lower() != "nan":
                    clean_numbers = re.findall(r'\d+\.?\d*', raw_val)
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
                            for part in [p.strip() for p in raw_val.split(",") if p.strip()]:
                                if part.lower() in lower_dict:
                                    if lower_dict[part.lower()]: final_values.add(lower_dict[part.lower()])
                                else: unmapped_values.add(f"Safilo -> {col_name}: '{part}'")
                        else: final_values.add(raw_val)
                return ", ".join(sorted(list(final_values)))

            elif col_name == "Glasses_lens_Colour" and mfg == "luxottica":
                if raw_val and raw_val.lower() != "nan":
                    matched = False
                    if col_name in VALUE_TRANSLATOR:
                        for keyword, mapped_val in VALUE_TRANSLATOR[col_name].items():
                            if keyword and keyword.lower() in raw_val.lower():
                                if mapped_val: final_values.add(mapped_val)
                                matched = True
                    if not matched: unmapped_values.add(f"Luxottica -> {col_name}: '{raw_val}'")
                return ", ".join(sorted(list(final_values)))

            elif raw_val and raw_val.lower() != "nan":
                if col_name in VALUE_TRANSLATOR:
                    translation_dict = VALUE_TRANSLATOR[col_name]
                    lower_dict = {str(k).lower(): v for k, v in translation_dict.items() if k}
                    for part in [p.strip() for p in raw_val.split(",") if p.strip()]:
                        if part.lower() in lower_dict:
                            if lower_dict[part.lower()]: final_values.add(lower_dict[part.lower()])
                        else: unmapped_values.add(f"{mfg.title()} -> {col_name}: '{part}'")
                else: final_values.add(raw_val)

            return ", ".join(sorted(list(final_values)))

        print(f"⚙️ Running custom rules engine for {mfg_name.title()}...")
        for target_col in new_df.columns:
            if target_col in VALUE_TRANSLATOR or target_col in ["Glasses_other_info", "Glasses_lens_effect", "SunGlasses_RX_lenses", "Glasses_type", "Glasses_shape", "Sunglasses_filter"]:
                new_df[target_col] = new_df.apply(lambda row: process_cell_strict(row, target_col, mfg_name), axis=1)

        # 4. ASSEMBLE NAMES
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
                    mc_parts = mat_num.split(" ")[0].replace("-", " ").split(" ")
                    model_out = mc_parts[0]
                    if len(mc_parts) > 1: color_out = mc_parts[1]
                    parts = [brand, mat_num.split(" ")[0].replace("-", " ")]
                else: parts = [brand]
            else:
                parts = [brand]
            return " ".join([p for p in parts if p]), model_out, color_out

        if not new_df.empty:
            temp_col = new_df.apply(lambda row: assemble_name_and_parts(row, mfg_name), axis=1)
            new_df["Assembled_Name"] = temp_col.apply(lambda x: x[0] if isinstance(x, (list, tuple)) else "")
            new_df["Extracted_Model"] = temp_col.apply(lambda x: x[1] if isinstance(x, (list, tuple)) else "")
            new_df["Extracted_Color"] = temp_col.apply(lambda x: x[2] if isinstance(x, (list, tuple)) else "")

        if "Manufacturer" in new_df.columns: new_df["Manufacturer"] = new_df["Manufacturer"].apply(lambda x: str(x).strip().title() if pd.notna(x) and str(x).strip().lower() not in ["nan", ""] else "")
        if "Brand" in new_df.columns: new_df["Brand"] = new_df["Brand"].apply(lambda x: str(x).strip().title() if pd.notna(x) and str(x).strip().lower() not in ["nan", ""] else "")

        # 5. ROUND DIMENSIONS
        for dim_col in ["Glasses_size_temple_length", "Glasses_size_lens_height", "Glasses_size_lens_width", "Glasses_size_bridge"]:
            if dim_col in new_df.columns:
                def round_dimension(val):
                    if pd.isna(val) or str(val).strip() == "" or str(val).strip().lower() == "nan": return ""
                    try:
                        clean_str = re.sub(r'[^\d,.-]', '', str(val).strip()).replace(',', '.')
                        if clean_str: return str(int(round(float(clean_str))))
                    except: pass
                    return str(val).strip()
                new_df[dim_col] = new_df[dim_col].apply(round_dimension)

        # 6. ZERO-STRIPPER & JOIN KEY
        if "Barcode" in new_df.columns:
            new_df["join_key"] = new_df["Barcode"].astype(str).str.strip().str.replace(r'\.0$', '', regex=True).str.lstrip('0')
            new_df = new_df[new_df["join_key"].notna() & (new_df["join_key"] != "nan") & (new_df["join_key"] != "")]
        else:
            print(f"❌ CRITICAL: 'Barcode' missing in {mfg_name} after extraction.")

        new_df["Producing_company"] = mfg_name.title()

        # Add to virtual catalog
        for brand in settings["brands"]:
            virtual_catalog[brand.lower().strip()] = new_df

    # --- FINAL COMPILATION ---
    print("\n🔗 Compiling all manufacturers into Master Database...")
    all_dfs = list(virtual_catalog.values())
    if not all_dfs:
        print("❌ No data processed. Exiting.")
        return

    master_db = pd.concat(all_dfs, ignore_index=True)
    master_db.drop_duplicates(subset=['join_key'], keep='first', inplace=True)
    
    # Put the Barcode/join_key cleanly into the database
    master_db.set_index('join_key', inplace=True)
    master_db_to_upload = master_db.reset_index()

    print(f"⬆️ Pushing clean Master Catalog ({len(master_db_to_upload)} rows) to Vault...")
    master_db_to_upload.to_sql(name='master_catalog', con=engine, if_exists='replace', index=False)
    
    print("\n✅ MASTER CATALOG SUCCESSFULLY BUILT AND STORED!")

    if unmapped_values:
        print("\n⚠️ UNMAPPED VALUES FOUND (You may want to add these to your dictionaries):")
        for val in sorted(unmapped_values):
            print(f"  - {val}")

if __name__ == "__main__":
    build_master_catalog()