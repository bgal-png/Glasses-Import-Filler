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
    KNOWN_BRANDS,
    classify_color,
)
from ingest import (
    load_single_catalog as _pure_load_single_catalog,
    perform_upsert as _pure_perform_upsert,
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
    return create_engine(DB_URL, pool_pre_ping=True, pool_recycle=300)

engine = get_engine()

# ==========================================
# 🧠 THE ENGINE (ADAPTED FOR UI)
# ==========================================
# All real processing logic lives in ingest.py so the headless GitHub Action
# (scripts/auto_ingest_safilo.py) uses the exact same code paths. These wrappers
# only add Streamlit UI calls (error toasts, unmapped-value expanders, brand-expansion).

def load_single_catalog(mfg_name, config_settings, file_path):
    """Streamlit wrapper around ingest.load_single_catalog."""
    try:
        processed_df, unmapped_values, skipped_not_mapped = _pure_load_single_catalog(
            mfg_name, config_settings, file_path
        )
    except ValueError as e:
        st.error(f"❌ {e}")
        return pd.DataFrame()

    # Surface diagnostic info via expandable UI sections
    if unmapped_values:
        with st.expander(f"⚠️ Unmapped Values Found in {mfg_name.title()} File"):
            for val in sorted(unmapped_values):
                st.write(f"- {val}")
    if skipped_not_mapped:
        with st.expander(f"ℹ️ Skipped 'NOT MAPPED' Values in {mfg_name.title()} File ({len(skipped_not_mapped)} unique)"):
            for val in sorted(skipped_not_mapped):
                st.write(f"- {val}")

    # Expand by brands defined in config (preserves prior behavior)
    if not processed_df.empty:
        all_brands_dfs = [processed_df.copy() for _ in config_settings["brands"]]
        return pd.concat(all_brands_dfs, ignore_index=True)
    return pd.DataFrame()


def perform_upsert(new_data_df):
    """Streamlit wrapper around ingest.perform_upsert."""
    msg = _pure_perform_upsert(new_data_df, engine)
    return msg


# --- Old inline implementation below kept as `_legacy_*` until removed in a follow-up commit ---
def _legacy_load_single_catalog(mfg_name, config_settings, file_path):
    """[DEAD CODE — replaced by ingest.load_single_catalog. Kept temporarily for safety.]"""
    unmapped_values = set()
    skipped_not_mapped = set()

    try:
        if file_path.endswith(".csv"):
            df = pd.read_csv(file_path, dtype=str, on_bad_lines="skip", sep=",")
            if len(df.columns) <= 1:
                # Single-column result means commas weren't the separator — retry with semicolons
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

    # 1b. Safilo: construct Combination from separate size fields (new CSV format)
    if mfg_name == "safilo" and all(c in df.columns for c in ["ItemSize", "ItemBridgeLength", "TempleLength"]):
        def _build_safilo_combination(row):
            parts = []
            for col in ["ItemSize", "ItemBridgeLength", "TempleLength"]:
                val = str(row.get(col, "")).strip()
                if val and val.lower() not in ("nan", ""):
                    try:
                        clean = re.sub(r"[^\d,.-]", "", val).replace(",", ".")
                        if clean:
                            val = str(int(round(float(clean))))
                    except Exception:
                        pass
                    parts.append(val)
            return "-".join(parts) if parts else ""
        new_df["Combination"] = df.apply(_build_safilo_combination, axis=1)

    # 2. Raw Clip-On Engine
    extracted_clip_ons = []
    clip_on_alerts = []

    for idx, raw_row in df.iterrows():
        clip_val = ""
        alert = False

        if mfg_name == "safilo":
            # New format: StyleD ending in "/C" marks the optical frame as bundled
            # with a magnetic clip-on. LenPolarized tells us if the clip is polarized.
            style_d = str(raw_row.get("StyleD", "")).strip().upper()
            pol = str(raw_row.get("LenPolarized", "")).strip().upper()
            prod_type = re.sub(r"\s+", " ", str(raw_row.get("TypeD", "")).strip().upper())

            is_clipon = (
                style_d.endswith("/C")
                or "CLIP-IN" in style_d
                or "CLIP-ON" in style_d
                or "CLIP ON" in style_d
                # Legacy format fallback (old Product Type Desc. column had "+ CLIP-ON")
                or "CLIP-ON" in prod_type
                or "CLIP ON" in prod_type
            )
            if is_clipon:
                if pol in ("X", "Y"): clip_val = "Magnetic sun clip-on p"
                else: clip_val = "Magnetic sun clip-on"

        elif mfg_name == "luxottica":
            # Find description column dynamically (encoding-safe)
            desc = ""
            for col_name in raw_row.index:
                if "popis" in col_name.lower() and "model" in col_name.lower():
                    desc = str(raw_row[col_name]).strip().upper()
                    break
            if desc == "CLIP ON":
                # Find polarized column dynamically
                pol = ""
                for col_name in raw_row.index:
                    if "polariz" in col_name.lower():
                        pol = str(raw_row[col_name]).strip().upper()
                        break
                if pol == "X":
                    clip_val = "Sun clip-on p"
                else:
                    clip_val = "Sun clip-on"

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

    # 2b. Luxottica Clip-On → Base Model Matching
    # Clip-ons are separate rows; attach their clip value to the base model (model without trailing "C")
    if mfg_name == "luxottica" and "Glasses_model" in new_df.columns:
        clip_mask = new_df["Extracted_Clip_on"].astype(str).str.strip().ne("") & new_df["Extracted_Clip_on"].notna()
        clip_rows = new_df[clip_mask]
        for idx, clip_row in clip_rows.iterrows():
            clip_model = str(clip_row.get("Glasses_model", "")).strip().lstrip("0")
            if clip_model.endswith("C"):
                base_model = clip_model[:-1]
                base_mask = new_df["Glasses_model"].astype(str).str.strip().str.lstrip("0") == base_model
                base_indices = new_df[base_mask & ~clip_mask].index
                for base_idx in base_indices:
                    if str(new_df.at[base_idx, "Extracted_Clip_on"]).strip() in ["", "nan"]:
                        new_df.at[base_idx, "Extracted_Clip_on"] = clip_row["Extracted_Clip_on"]
                        new_df.at[base_idx, "Clip_on_Alert"] = clip_row["Clip_on_Alert"]
                        # Copy clip-on lens colour to base model (classify raw value)
                        clip_lens_colour = str(clip_row.get("Glasses_lens_Colour", "")).strip()
                        if clip_lens_colour and clip_lens_colour.lower() != "nan":
                            classified = classify_color(clip_lens_colour, "lens")
                            new_df.at[base_idx, "Clip_on_lens_colour"] = classified if classified else clip_lens_colour

    # 3. Custom Rules Strict Engine
    def process_cell_strict(row, col_name, mfg):
        final_values = set()
        raw_val = str(row.get(col_name, "")).strip()

        if col_name == "Glasses_other_info":
            if mfg == "safilo":
                # Flex: prefer Hinge_raw column (new CSV format), fall back to model name (old format)
                hinge = str(row.get("Hinge_raw", "")).strip().upper()
                if hinge and hinge not in ("NAN", "NO FLEX", ""):
                    if "FLEX" in hinge:
                        final_values.add("Flex")
                elif pd.notna(row.get("Glasses_model")) and "FLEX" in str(row["Glasses_model"]).upper():
                    final_values.add("Flex")
                # Double bridge from Shape value
                shape_raw = str(raw_val).strip().upper()
                if "DOUBLE BRIDGE" in shape_raw:
                    final_values.add("Double bridge")
            elif mfg == "luxottica":
                raw_info = str(row.get("Glasses_other_info", "")).strip().upper()
                if raw_info == "X": final_values.add("Flex")
                if pd.notna(row.get("Glasses_collection")) and str(row["Glasses_collection"]).strip().upper() == "X": final_values.add("Flexible glasses")
            elif mfg in ["kering", "marcolin"]:
                if pd.notna(row.get("Family_descriptions_raw")):
                    if "double bridge" in str(row["Family_descriptions_raw"]).lower(): final_values.add("Double bridge")

        elif col_name == "Glasses_lens_effect":
            if mfg == "safilo":
                # New format uses Y/N; old format used X/0 — handle both
                pol = str(row.get("Polarized_raw", "")).strip().upper()
                if pol in ("X", "Y"): final_values.add("Polarized")
                phot = str(row.get("Photochromic_raw", "")).strip().upper()
                if phot in ("X", "Y"): final_values.add("Photochromic")
                # Treatement Description column removed in new CSV format
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

        elif col_name == "SunGlasses_RX_lenses":
            raw_rx = str(row.get(col_name, "")).strip().upper()
            if mfg in ["safilo", "kering", "marcolin"]:
                if raw_rx in ("X", "Y"): final_values.add("Yes")
            elif mfg == "luxottica":
                if raw_rx == "YES": final_values.add("Yes")

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
                    elif 0 < vlt < 8: final_values.add("Category 4"); matched_by_math = True
                    # vlt == 0 is treated as "no filter data" (optical frame) — no category assigned

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

        elif col_name in ("Glasses_lens_Colour", "Frame_Colour", "Temple_Colour"):
            if raw_val and raw_val.lower() != "nan":
                color_type = "lens" if col_name == "Glasses_lens_Colour" else "frame"
                result = classify_color(raw_val, color_type)
                if result:
                    for c in result.split("|"):
                        final_values.add(c)
                else:
                    unmapped_values.add(f"{mfg.title()} -> {col_name}: '{raw_val}'")

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

        # Filter out "NOT MAPPED" values — keep cell clean, track for summary
        if "NOT MAPPED" in final_values:
            skipped_not_mapped.add(f"{mfg.title()} -> {col_name}")
            final_values.discard("NOT MAPPED")
        return "|".join(sorted(final_values))

    for target_col in new_df.columns:
        if target_col in VALUE_TRANSLATOR or target_col in ["Glasses_other_info", "Glasses_lens_effect", "SunGlasses_RX_lenses", "Glasses_type", "Glasses_shape", "Sunglasses_filter", "Glasses_lens_Colour", "Frame_Colour", "Temple_Colour"]:
            new_df[target_col] = new_df.apply(lambda row: process_cell_strict(row, target_col, mfg_name), axis=1)

    # ==========================================
    # 🧼 PRE-CLEAN: EXTRACT "KIDS" FROM BRAND/MFG
    # ==========================================
    if "Brand" not in new_df.columns: new_df["Brand"] = ""
    if "Manufacturer" not in new_df.columns: new_df["Manufacturer"] = ""

    # Flag the row if "Kids" is anywhere in the brand or manufacturer
    new_df["Is_Kids"] = new_df["Brand"].astype(str).str.contains(r"(?i)\bkids\b", regex=True, na=False) | \
                        new_df["Manufacturer"].astype(str).str.contains(r"(?i)\bkids\b", regex=True, na=False)

    # Scrub "Kids" completely out of those columns and fix the spacing
    new_df["Brand"] = new_df["Brand"].astype(str).str.replace(r"(?i)\bkids\b", "", regex=True).str.replace(r"\s+", " ", regex=True).str.strip()
    new_df["Brand"] = new_df["Brand"].apply(lambda x: str(x).title() if x and x.lower() != "nan" else "")

    new_df["Manufacturer"] = new_df["Manufacturer"].astype(str).str.replace(r"(?i)\bkids\b", "", regex=True).str.replace(r"\s+", " ", regex=True).str.strip()
    new_df["Manufacturer"] = new_df["Manufacturer"].apply(lambda x: str(x).title() if x and x.lower() != "nan" else "")

    # ==========================================
    # 🏷️ WHITELIST BRAND CLEANING
    # ==========================================
    # Match raw brand to closest known brand from MANUFACTURER_CONFIG.
    # Handles cases like "Polaroid Kids" → "Polaroid", "Ray-Ban Vista" → "Ray-Ban"
    # by checking if the raw value starts with a known brand name.
    _brand_lookup = sorted(KNOWN_BRANDS, key=len, reverse=True)  # longest first

    def _clean_brand_to_whitelist(raw):
        raw = str(raw).strip()
        if not raw or raw.lower() == "nan":
            return raw
        raw_lower = raw.lower()
        for known in _brand_lookup:
            if raw_lower == known.lower():
                return known  # exact match
            if raw_lower.startswith(known.lower() + " "):
                return known  # e.g. "Polaroid Kids" → "Polaroid"
        return raw  # no match — keep original

    # Brand name corrections (word swaps, special cases) — run BEFORE whitelist
    # so "Moschino Love" → "Love Moschino" before the whitelist can collapse it to "Moschino"
    BRAND_CORRECTIONS = {
        "moschino love": "Love Moschino",
        "prive' revaux": "Prive Revaux",
    }
    def _correct_brand(raw):
        raw = str(raw).strip()
        if raw.lower() in BRAND_CORRECTIONS:
            return BRAND_CORRECTIONS[raw.lower()]
        return raw

    new_df["Brand"] = new_df["Brand"].apply(_correct_brand)
    new_df["Manufacturer"] = new_df["Manufacturer"].apply(_correct_brand)

    new_df["Brand"] = new_df["Brand"].apply(_clean_brand_to_whitelist)
    new_df["Manufacturer"] = new_df["Manufacturer"].apply(_clean_brand_to_whitelist)

    # ==========================================
    # 🏗️ ASSEMBLE MODEL AND NAMES
    # ==========================================
    def assemble_name_and_parts(row, mfg):
        brand = str(row.get("Brand", "")).strip()
        is_kids = row.get("Is_Kids", False)
        
        model_out, color_out = "", ""

        if mfg == "safilo":
            model_out = str(row.get("Glasses_model", "")).strip()
            color_out = str(row.get("Glasses_color_code", "")).strip()
            if model_out.lower() == "nan": model_out = ""
            if color_out.lower() == "nan": color_out = ""
            
            if is_kids and model_out: model_out = f"Kids {model_out}"
            elif is_kids: model_out = "Kids"
            
            parts = [brand, model_out, color_out]
            
        elif mfg == "luxottica":
            model_out = str(row.get("Glasses_model", "")).strip().lstrip("0")
            color_out = str(row.get("Glasses_color_code", "")).strip()
            if model_out.lower() == "nan": model_out = ""
            if color_out.lower() == "nan": color_out = ""
            
            if is_kids and model_out: model_out = f"Kids {model_out}"
            elif is_kids: model_out = "Kids"
            
            parts = [brand, model_out, color_out]
            
        elif mfg in ["kering", "marcolin"]:
            mat_num = str(row.get("Material_Number", "")).strip()
            if mat_num and mat_num.lower() != "nan":
                first_part = mat_num.split(" ")[0]
                model_color = first_part.replace("-", " ")
                mc_parts = model_color.split(" ")
                
                model_out = mc_parts[0]
                if is_kids and model_out: model_out = f"Kids {model_out}"
                elif is_kids: model_out = "Kids"
                
                if len(mc_parts) > 1: 
                    color_out = mc_parts[1]
                    parts = [brand, model_out, color_out]
                else: 
                    parts = [brand, model_out]
            else: 
                if is_kids: model_out = "Kids"
                parts = [brand, model_out] if model_out else [brand]
        else: 
            if is_kids: model_out = "Kids"
            parts = [brand, model_out] if model_out else [brand]

        final_name = " ".join([p for p in parts if p])
        return final_name, model_out, color_out

    if not new_df.empty:
        temp_col = new_df.apply(lambda row: assemble_name_and_parts(row, mfg_name), axis=1)
        new_df["Assembled_Name"] = temp_col.apply(lambda x: x[0] if isinstance(x, (list, tuple)) else "")
        new_df["Extracted_Model"] = temp_col.apply(lambda x: x[1] if isinstance(x, (list, tuple)) else "")
        new_df["Extracted_Color"] = temp_col.apply(lambda x: x[2] if isinstance(x, (list, tuple)) else "")
        
    if "Is_Kids" in new_df.columns:
        new_df.drop(columns=["Is_Kids"], inplace=True)

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

    if unmapped_values or skipped_not_mapped:
        if unmapped_values:
            with st.expander(f"⚠️ Unmapped Values Found in {mfg_name.title()} File"):
                for val in sorted(unmapped_values): st.write(f"- {val}")
        if skipped_not_mapped:
            with st.expander(f"ℹ️ Skipped 'NOT MAPPED' Values in {mfg_name.title()} File ({len(skipped_not_mapped)} unique)"):
                for val in sorted(skipped_not_mapped): st.write(f"- {val}")

    if all_brands_dfs:
        combined_df = pd.concat(all_brands_dfs, ignore_index=True)
        return combined_df
    return pd.DataFrame()

# ==========================================
# 🔄 UPSERT ENGINE (WITH DETAILED TRACKING)
# ==========================================
def _legacy_perform_upsert(new_data_df):
    """[DEAD CODE — replaced by ingest.perform_upsert. Kept temporarily for safety.]"""
    new_data_df.drop_duplicates(subset=["join_key"], keep="last", inplace=True)
    new_data_df.set_index("join_key", inplace=True)
    
    try:
        existing_df = pd.read_sql_table('master_catalog', con=engine)
        existing_df.set_index("join_key", inplace=True)
        
        # 📊 TRACKING: Calculate exactly what overlaps and what is new
        common_indices = new_data_df.index.intersection(existing_df.index)
        updated_count = len(common_indices)
        
        # A. Update existing rows with fresh data
        existing_df.update(new_data_df)
        
        # B. Find brand new rows that don't exist in the database yet
        new_rows = new_data_df[~new_data_df.index.isin(existing_df.index)]
        added_count = len(new_rows)
        
        # C. Stitch together
        final_df = pd.concat([existing_df, new_rows])
        
        upsert_msg = f"🔄 Refreshed {updated_count} existing products. ✨ Added {added_count} completely new products!"
        
    except Exception as e:
        final_df = new_data_df
        upsert_msg = f"✨ Created database from scratch with {len(new_data_df)} products!"

    # Push back to Supabase
    final_df.reset_index().to_sql('master_catalog', con=engine, if_exists='replace', index=False, )
    return upsert_msg

# ==========================================
# 🖥️ USER INTERFACE
# ==========================================

# --- 📊 DATABASE STATUS ---
st.divider()
try:
    db_count_df = pd.read_sql("SELECT COUNT(*) as total FROM master_catalog", con=engine)
    total_items = db_count_df["total"].iloc[0]
    st.metric("📦 Total Items in Database", f"{total_items:,}")
except:
    st.info("📦 Database is empty or not yet created.")

# --- 🔍 BARCODE SEARCH & EDIT ---
with st.expander("🔍 Barcode Lookup & Editor", expanded=False):
    search_col1, search_col2 = st.columns([3, 1])
    with search_col1:
        search_ean = st.text_input("Enter EAN / Barcode to search:", placeholder="e.g. 8056597123456")
    with search_col2:
        st.write("")
        st.write("")
        search_btn = st.button("Search", use_container_width=True)

    if search_btn and search_ean:
        st.session_state["lookup_ean"] = search_ean

    if "lookup_ean" in st.session_state and st.session_state["lookup_ean"]:
        clean_search = re.sub(r"\.0$", "", str(st.session_state["lookup_ean"]).strip()).lstrip("0")
        try:
            result_df = pd.read_sql_table('master_catalog', con=engine)
            if 'join_key' in result_df.columns:
                result_df['join_key'] = result_df['join_key'].astype(str).str.strip()
                match = result_df[result_df['join_key'] == clean_search]
                if not match.empty:
                    st.success(f"✅ Barcode '{st.session_state['lookup_ean']}' found!")

                    row_data = match.iloc[0].fillna("")
                    row_data = row_data.apply(lambda x: str(x).strip() if str(x).strip().lower() != "nan" else "")

                    # Editable fields
                    view_tab, edit_tab = st.tabs(["📋 View", "✏️ Edit"])

                    with view_tab:
                        st.dataframe(row_data.to_frame("Value"), use_container_width=True)

                    with edit_tab:
                        st.caption("Edit any field below and click Save to update the database.")
                        edited_values = {}
                        cols_to_edit = [c for c in match.columns if c != "join_key"]

                        for col in cols_to_edit:
                            current_val = str(row_data.get(col, "")).strip()
                            if current_val.lower() == "nan":
                                current_val = ""
                            edited_values[col] = st.text_input(
                                col,
                                value=current_val,
                                key=f"edit_{col}_{clean_search}"
                            )

                        if st.button("💾 Save Changes", type="primary"):
                            try:
                                # Build update only for changed values
                                changes = {}
                                for col in cols_to_edit:
                                    old_val = str(row_data.get(col, "")).strip()
                                    if old_val.lower() == "nan":
                                        old_val = ""
                                    new_val = edited_values[col].strip()
                                    if new_val != old_val:
                                        changes[col] = new_val

                                if changes:
                                    # Update in database
                                    idx_pos = result_df.index[result_df['join_key'] == clean_search]
                                    for col, val in changes.items():
                                        result_df.loc[idx_pos, col] = val
                                    result_df.to_sql('master_catalog', con=engine, if_exists='replace', index=False, )
                                    st.success(f"✅ Updated {len(changes)} field(s): {', '.join(changes.keys())}")
                                    st.rerun()
                                else:
                                    st.info("No changes detected.")
                            except Exception as e:
                                st.error(f"Failed to save: {e}")
                else:
                    st.error(f"❌ Barcode '{st.session_state['lookup_ean']}' (cleaned: {clean_search}) not found.")
        except Exception as e:
            st.error(f"Search failed: {e}")

st.divider()

col1, col2 = st.columns(2)

# --- 🏭 COLUMN 1: MANUFACTURER CATALOGS ---
with col1:
    st.subheader("🏭 Update Manufacturer Catalogs")
    st.write("Upload a raw manufacturer file. It will be cleaned, processed, and merged into the Master Vault.")
    
    mfg_choice = st.selectbox("Select Manufacturer:", ["safilo", "luxottica", "marcolin", "kering", "derigo", "thelios"])
    uploaded_mfg = st.file_uploader(f"Upload new {mfg_choice.title()} file", type=["csv", "xlsx"])
    
    if uploaded_mfg and st.button(f"🚀 Process & Upsert {mfg_choice.title()} Catalog", type="primary"):
        with st.spinner(f"Processing raw {mfg_choice.title()} file through the Rules Engine..."):
            
            temp_path = f"temp_{uploaded_mfg.name}"
            processed_df = pd.DataFrame()
            
            try:
                # Save uploaded file temporarily to disk securely
                with open(temp_path, "wb") as f:
                    f.write(uploaded_mfg.getbuffer())
                    
                config_settings = MANUFACTURER_CONFIG[mfg_choice]
                processed_df = load_single_catalog(mfg_choice, config_settings, temp_path)
                
            finally:
                # 🧹 SAFELY clean up the temp file (prevents FileNotFoundError crashes)
                if os.path.exists(temp_path):
                    os.remove(temp_path)
            
            if not processed_df.empty:
                st.success(f"🧹 Cleaned successfully! Extracted {len(processed_df)} raw rows. Pushing to Cloud...")
                msg = perform_upsert(processed_df)
                st.success(msg)
            else:
                st.error("Failed to extract any data. Please check the file format.")

    st.divider()
    st.markdown("**⚠️ Danger zone — delete all rows for a manufacturer**")
    st.caption("Use this before re-uploading after a structure change (e.g. Safilo CSV migration). Rows are matched by `Producing_company`.")
    del_mfg = st.selectbox(
        "Manufacturer to wipe:",
        ["safilo", "luxottica", "marcolin", "kering", "derigo", "thelios"],
        key="del_mfg_choice",
    )
    confirm_text = st.text_input(
        f"Type **DELETE {del_mfg.upper()}** to confirm:",
        key="del_confirm_input",
        placeholder=f"DELETE {del_mfg.upper()}",
    )
    if st.button(f"🗑️ Delete all {del_mfg.title()} rows", type="secondary"):
        if confirm_text.strip() != f"DELETE {del_mfg.upper()}":
            st.error("Confirmation text doesn't match. Nothing deleted.")
        else:
            with st.spinner(f"Loading master_catalog to filter out {del_mfg.title()}..."):
                try:
                    full_df = pd.read_sql_table("master_catalog", con=engine)
                    before_count = len(full_df)
                    if "Producing_company" not in full_df.columns:
                        st.error("⚠️ master_catalog has no `Producing_company` column — cannot filter.")
                    else:
                        target = del_mfg.title()
                        keep_df = full_df[full_df["Producing_company"].astype(str).str.strip().str.lower() != target.lower()]
                        deleted = before_count - len(keep_df)
                        if deleted == 0:
                            st.info(f"No rows found with Producing_company = '{target}'. Nothing to delete.")
                        else:
                            keep_df.to_sql("master_catalog", con=engine, if_exists="replace", index=False)
                            st.success(f"✅ Deleted {deleted:,} {target} rows. {len(keep_df):,} rows remain.")
                except Exception as e:
                    st.error(f"Delete failed: {e}")

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
            df_pkg.to_sql('package_data', engine, if_exists='replace', index=False, )
            st.success(f"✅ Package Data updated! ({len(df_pkg)} items)")
            
    st.divider()
    
    # Global Categories
    hist_file = st.file_uploader("Upload new `global_categories.xlsx`", type=["xlsx"])
    if hist_file and st.button("⬆️ Replace Global Categories in Cloud"):
        with st.spinner("Uploading Global Categories..."):
            df_hist = pd.read_excel(hist_file, dtype=str, engine="openpyxl")
            if "Items type" in df_hist.columns:
                df_glasses = df_hist[df_hist["Items type"].astype(str).str.strip().str.lower() == "glasses"]
                df_glasses.columns = df_glasses.columns.astype(str).str.strip()
                # Only keep columns we actually use to save storage
                keep_cols = [c for c in ["Brand", "Glasses contain"] if c in df_glasses.columns]
                df_glasses = df_glasses[keep_cols]
                df_glasses.to_sql('historical_data', engine, if_exists='replace', index=False, )
                st.success(f"✅ Global Categories updated! ({len(df_glasses)} glasses mapped)")
            else:
                st.error("⚠️ 'Items type' column missing. Could not filter/upload historical data.")

    st.divider()

    # Item Origin
    origin_file = st.file_uploader("Upload new `item_origin.xlsx`", type=["xlsx"])
    if origin_file and st.button("⬆️ Replace Item Origin in Cloud"):
        with st.spinner("Uploading Item Origin..."):
            df_origin = pd.read_excel(origin_file, dtype=str, engine="openpyxl")
            df_origin.columns = df_origin.columns.astype(str).str.strip()
            keep_cols = [c for c in ["item_name", "country_master"] if c in df_origin.columns]
            if len(keep_cols) == 2:
                df_origin = df_origin[keep_cols]
                df_origin = df_origin.dropna(subset=["item_name", "country_master"])
                df_origin.to_sql('origin_data', engine, if_exists='replace', index=False, )
                st.success(f"✅ Item Origin updated! ({len(df_origin)} items)")
            else:
                st.error("⚠️ 'item_name' or 'country_master' column missing.")

# ==========================================
# 📒 CREATED-ITEMS REGISTRY (history of what we've already made)
# ==========================================
st.divider()
st.subheader("📒 Created Items Registry")
st.caption(
    "Keep a record of products you've already created. Store past filled files "
    "(Name + barcode + size), then later check a list of barcodes to see which "
    "you've already done."
)


def _clean_bc(x):
    """Normalize a barcode the same way as master_catalog join_key."""
    return re.sub(r"\.0$", "", str(x).strip()).lstrip("0")


def _read_any(uploaded):
    """Read an uploaded xlsx/csv into a string DataFrame with stripped headers."""
    if uploaded.name.lower().endswith(".csv"):
        try:
            d = pd.read_csv(uploaded, dtype=str, sep=",", on_bad_lines="skip")
            if len(d.columns) <= 1:
                uploaded.seek(0)
                d = pd.read_csv(uploaded, dtype=str, sep=";", on_bad_lines="skip")
        except Exception:
            uploaded.seek(0)
            d = pd.read_csv(uploaded, dtype=str, sep=";", on_bad_lines="skip")
    else:
        d = pd.read_excel(uploaded, dtype=str, engine="openpyxl")
    d.columns = (
        d.columns.astype(str)
        .str.replace(r"[\r\n\t]", " ", regex=True)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )
    return d


def _find_col(df, candidates):
    """Return the first column whose name matches any candidate (case-insensitive)."""
    lower_map = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in lower_map:
            return lower_map[cand.lower()]
    return None


reg_col1, reg_col2 = st.columns(2)

# --- STORE side ---
with reg_col1:
    st.markdown("**➕ Store created items**")
    st.caption("Upload one or more older filled files. Stores Name + barcode + size, merging by barcode.")
    stored_files = st.file_uploader(
        "Upload filled file(s)", type=["xlsx", "csv"], accept_multiple_files=True, key="reg_store"
    )
    if stored_files and st.button("💾 Store as created items", type="primary"):
        with st.spinner("Storing..."):
            new_rows = []
            problems = []
            for uf in stored_files:
                try:
                    d = _read_any(uf)
                except Exception as e:
                    problems.append(f"{uf.name}: could not read ({e})")
                    continue
                bc_col = _find_col(d, ["Barcode", "EAN", "UPC", "EAN/UPC"])
                name_col = _find_col(d, ["Glasses name", "XML description", "Name"])
                size_col = _find_col(d, ["Combination (size on glasses)", "Combination", "Size"])
                if not bc_col:
                    problems.append(f"{uf.name}: no Barcode column found")
                    continue
                tmp = pd.DataFrame()
                tmp["join_key"] = d[bc_col].apply(_clean_bc)
                tmp["barcode"] = d[bc_col].astype(str).str.strip()
                tmp["name"] = d[name_col].astype(str).str.strip() if name_col else ""
                tmp["size"] = d[size_col].astype(str).str.strip() if size_col else ""
                tmp = tmp[tmp["join_key"].notna() & (tmp["join_key"] != "") & (tmp["join_key"] != "nan")]
                new_rows.append(tmp)

            if problems:
                for p in problems:
                    st.warning(f"⚠️ {p}")

            if new_rows:
                incoming = pd.concat(new_rows, ignore_index=True)
                try:
                    existing = pd.read_sql_table("created_items", con=engine)
                except Exception:
                    existing = pd.DataFrame(columns=["join_key", "barcode", "name", "size"])
                before = len(existing)
                combined = pd.concat([existing, incoming], ignore_index=True)
                combined.drop_duplicates(subset=["join_key"], keep="last", inplace=True)
                combined.to_sql("created_items", con=engine, if_exists="replace", index=False)
                added = len(combined) - before
                st.success(
                    f"✅ Stored {len(incoming)} rows from {len(new_rows)} file(s). "
                    f"Registry now holds {len(combined):,} unique items ({added:,} new)."
                )
            else:
                st.error("No usable rows found in the uploaded file(s).")

# --- CHECK side ---
with reg_col2:
    st.markdown("**🔎 Check barcodes against registry**")
    st.caption("Upload a file with barcodes to see which you've already created.")
    check_file = st.file_uploader("Upload barcode list", type=["xlsx", "csv"], key="reg_check")
    if check_file and st.button("🔎 Check barcodes"):
        with st.spinner("Checking..."):
            try:
                registry = pd.read_sql_table("created_items", con=engine)
            except Exception:
                registry = pd.DataFrame(columns=["join_key", "barcode", "name", "size"])

            if registry.empty:
                st.warning("⚠️ Registry is empty — store some filled files first.")
            else:
                d = _read_any(check_file)
                bc_col = _find_col(d, ["Barcode", "EAN", "UPC", "EAN/UPC"]) or d.columns[0]
                reg_lookup = registry.set_index("join_key")
                results = []
                for raw in d[bc_col]:
                    key = _clean_bc(raw)
                    if not key or key == "nan":
                        continue
                    if key in reg_lookup.index:
                        row = reg_lookup.loc[key]
                        if isinstance(row, pd.DataFrame):
                            row = row.iloc[0]
                        results.append({
                            "Barcode": str(raw).strip(),
                            "Status": "✅ Already created",
                            "Name": row.get("name", ""),
                            "Size": row.get("size", ""),
                        })
                    else:
                        results.append({
                            "Barcode": str(raw).strip(),
                            "Status": "🆕 New",
                            "Name": "",
                            "Size": "",
                        })
                res_df = pd.DataFrame(results)
                already = (res_df["Status"] == "✅ Already created").sum()
                new_n = (res_df["Status"] == "🆕 New").sum()
                m1, m2 = st.columns(2)
                m1.metric("Already created", f"{already}")
                m2.metric("New", f"{new_n}")
                st.dataframe(res_df, use_container_width=True, hide_index=True)
                st.download_button(
                    "📥 Download result (CSV)",
                    data=res_df.to_csv(index=False).encode("utf-8-sig"),
                    file_name="barcode_check_result.csv",
                    mime="text/csv",
                )

# ==========================================
# 🎨 FILL MISSING COLOURS FROM PHOTOS
# ==========================================
import io
import zipfile
from dictionaries import _FRAME_TEMPLE_KEYWORDS, _LENS_KEYWORDS

st.divider()
st.subheader("🎨 Fill Missing Colours from Photos")
st.caption(
    "Upload a ZIP of product photos (filenames must contain the model + colour "
    "code). The tool finds catalogue items with missing colours, matches a photo "
    "to each, and lets you assign colours with one click — writing to every "
    "barcode that shares that model + colour."
)

# Canonical colour categories (distinct system colours from the classifier)
_FRAME_COLOURS = list(dict.fromkeys(v for _, v in _FRAME_TEMPLE_KEYWORDS))
_LENS_COLOURS = list(dict.fromkeys(v for _, v in _LENS_KEYWORDS))

_COLOUR_FIELDS = [
    # (master_catalog column, human label, which palette, condition)
    ("Frame_Colour", "Frame colour", "frame", "any"),
    ("Temple_Colour", "Temple colour", "frame", "any"),
    ("Glasses_lens_Colour", "Lens colour", "lens", "sunglasses"),
    ("Clip_on_lens_colour", "Clip-on lens colour", "lens", "clip"),
]


def _norm(s):
    return re.sub(r"[^A-Za-z0-9]", "", str(s or "")).upper()


with st.expander("🎨 Open the colour-filling tool", expanded=False):
    # --- Scope + upload ---
    try:
        _companies_df = pd.read_sql('SELECT DISTINCT "Producing_company" FROM master_catalog', con=engine)
        all_companies = sorted(
            c for c in _companies_df["Producing_company"].dropna().astype(str).str.strip().unique() if c
        )
    except Exception:
        all_companies = []
    scope = st.multiselect(
        "Limit to manufacturers (optional — leave empty to scan ALL):",
        all_companies,
        default=[],
        key="colfill_scope",
    )
    photos_zip = st.file_uploader("Upload ZIP of photos", type=["zip"], key="colfill_zip")

    if photos_zip and st.button("🔍 Build worklist", key="colfill_build"):
        with st.spinner("Reading photos and scanning catalogue for missing colours..."):
            zbytes = photos_zip.getvalue()
            # Index photos by normalized basename
            photo_index = []  # (normalized_name, original_name)
            try:
                with zipfile.ZipFile(io.BytesIO(zbytes)) as zf:
                    for n in zf.namelist():
                        if n.endswith("/"):
                            continue
                        base = n.rsplit("/", 1)[-1]
                        if base.lower().rsplit(".", 1)[-1] in ("jpg", "jpeg", "png", "webp", "gif"):
                            stem = base.rsplit(".", 1)[0]
                            photo_index.append((_norm(stem), n))
            except Exception as e:
                st.error(f"Could not read ZIP: {e}")
                photo_index = []

            mc = pd.read_sql_table("master_catalog", con=engine)
            if scope and "Producing_company" in mc.columns:
                mc = mc[mc["Producing_company"].astype(str).isin(scope)]

            def _blank(v):
                return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() in ("", "nan")

            # Group rows by (model, colour) — colour is shared across sizes
            groups = {}
            for _, row in mc.iterrows():
                model = str(row.get("Extracted_Model", "")).strip()
                colour = str(row.get("Glasses_color_code", "")).strip()
                if not model or model.lower() == "nan":
                    continue
                g_type = str(row.get("Glasses_type", "")).strip()
                has_clip = not _blank(row.get("Extracted_Clip_on", ""))
                missing = []
                for col, label, palette, cond in _COLOUR_FIELDS:
                    if not _blank(row.get(col, "")):
                        continue
                    if cond == "sunglasses" and "Sunglasses" not in g_type:
                        continue
                    if cond == "clip" and not has_clip:
                        continue
                    missing.append(col)
                # Sunglasses whose lens effect doesn't already include Gradient
                # can be marked as gradient during review.
                lens_effect = str(row.get("Glasses_lens_effect", "")).strip()
                can_gradient = ("Sunglasses" in g_type) and ("gradient" not in lens_effect.lower())
                if not missing and not can_gradient:
                    continue
                key = (_norm(model), _norm(colour))
                if key not in groups:
                    groups[key] = {
                        "brand": str(row.get("Brand", "")).strip(),
                        "model": model,
                        "colour_code": colour,
                        "size": str(row.get("Combination", "")).strip(),
                        "type": g_type,
                        "name": str(row.get("Assembled_Name", "")).strip(),
                        "barcodes": set(),
                        "missing": set(),
                        "can_gradient": False,
                    }
                groups[key]["barcodes"].add(str(row.get("join_key", "")).strip())
                groups[key]["missing"].update(missing)
                if can_gradient:
                    groups[key]["can_gradient"] = True

            # Match each group to a photo
            worklist = []
            unmatched = 0
            for (model_n, colour_n), g in groups.items():
                size_n = _norm(g["size"])
                sizecolour_n = size_n + colour_n
                found = None
                for norm_name, orig in photo_index:
                    if model_n and model_n in norm_name and (
                        (colour_n and colour_n in norm_name) or (sizecolour_n and sizecolour_n in norm_name)
                    ):
                        found = orig
                        break
                if found:
                    worklist.append({
                        "key": f"{model_n}|{colour_n}",
                        "photo": found,
                        "brand": g["brand"], "model": g["model"], "colour_code": g["colour_code"],
                        "type": g["type"], "name": g["name"],
                        "barcodes": sorted(g["barcodes"]),
                        "missing": [c for c, _, _, _ in _COLOUR_FIELDS if c in g["missing"]],
                        "can_gradient": g["can_gradient"],
                    })
                else:
                    unmatched += 1

            st.session_state["colfill_zip_bytes"] = zbytes
            st.session_state["colfill_worklist"] = worklist
            st.session_state["colfill_idx"] = 0
            st.session_state["colfill_assign"] = {}
            st.success(
                f"Found {len(groups)} model/colour groups with missing colours. "
                f"Matched a photo for {len(worklist)}; {unmatched} had no matching photo."
            )

    # --- Review UI ---
    worklist = st.session_state.get("colfill_worklist")
    if worklist:
        idx = st.session_state.get("colfill_idx", 0)
        assign = st.session_state.setdefault("colfill_assign", {})
        total = len(worklist)
        done = len(assign)

        label_map = {c: lbl for c, lbl, _, _ in _COLOUR_FIELDS}
        palette_map = {c: (_FRAME_COLOURS if p == "frame" else _LENS_COLOURS) for c, _, p, _ in _COLOUR_FIELDS}
        st.write(f"**{total} items to review**  ·  {done} assigned so far")

        layout = st.radio("Layout", ["🔲 Grid (fast)", "🖼️ One-by-one"], horizontal=True, key="colfill_layout")

        if layout == "🔲 Grid (fast)":
            csel1, csel2 = st.columns([1, 2])
            per_page = csel1.selectbox("Per page", [8, 12, 20, 30], index=1, key="colfill_perpage")
            n_pages = max(1, (total + per_page - 1) // per_page)
            page = csel2.number_input(f"Page (1–{n_pages})", 1, n_pages, 1, key="colfill_gridpage") - 1
            start, end = page * per_page, min(page * per_page + per_page, total)
            page_items = worklist[start:end]

            st.caption(f"Showing {start + 1}–{end} of {total}. Set the dropdowns, then **Save this page** "
                       "(changing dropdowns won't reload — only Save does). Save each page before paging.")

            with st.form(f"colfill_grid_{page}"):
                ncol = 4
                try:
                    zf = zipfile.ZipFile(io.BytesIO(st.session_state["colfill_zip_bytes"]))
                except Exception:
                    zf = None
                for r in range(0, len(page_items), ncol):
                    cols = st.columns(ncol)
                    for j, item in enumerate(page_items[r:r + ncol]):
                        with cols[j]:
                            if zf is not None:
                                try:
                                    st.image(zf.read(item["photo"]), use_container_width=True)
                                except Exception:
                                    st.write("(photo error)")
                            st.caption(f"{item['brand']} {item['model']} {item['colour_code']}")
                            cur = assign.get(item["key"], {})
                            for field in item["missing"]:
                                palette = palette_map[field]
                                opts = ["—"] + palette
                                chosen = cur.get(field)
                                st.selectbox(
                                    label_map[field], opts,
                                    index=(opts.index(chosen) if chosen in opts else 0),
                                    key=f"grid_{item['key']}_{field}",
                                )
                            if item.get("can_gradient"):
                                st.checkbox("Gradient lens", value=bool(cur.get("__gradient__")),
                                            key=f"grid_{item['key']}__gradient__")
                saved = st.form_submit_button("💾 Save this page")

            if saved:
                for item in page_items:
                    for field in item["missing"]:
                        val = st.session_state.get(f"grid_{item['key']}_{field}")
                        if val and val != "—":
                            assign.setdefault(item["key"], {})[field] = val
                        elif item["key"] in assign and field in assign[item["key"]]:
                            del assign[item["key"]][field]
                    if item.get("can_gradient"):
                        if st.session_state.get(f"grid_{item['key']}__gradient__"):
                            assign.setdefault(item["key"], {})["__gradient__"] = True
                        elif item["key"] in assign and "__gradient__" in assign[item["key"]]:
                            del assign[item["key"]]["__gradient__"]
                    if item["key"] in assign and not assign[item["key"]]:
                        del assign[item["key"]]
                st.session_state["colfill_assign"] = assign
                st.success(f"Saved page {page + 1}. {len(assign)} groups assigned so far.")
                st.rerun()

        else:
            idx = min(idx, total)
            st.progress(min(idx, total) / total if total else 0.0)
            if idx >= total:
                st.success("🎉 Reached the end of the worklist.")
            else:
                item = worklist[idx]
                key = item["key"]
                st.write(f"**Item {idx + 1} / {total}**")
                colL, colR = st.columns([1, 1])
                with colL:
                    try:
                        with zipfile.ZipFile(io.BytesIO(st.session_state["colfill_zip_bytes"])) as zf:
                            st.image(zf.read(item["photo"]), use_container_width=True)
                    except Exception:
                        st.warning("Could not display photo.")
                    st.caption(item["photo"])
                with colR:
                    st.markdown(f"**{item['name'] or item['brand']}**")
                    st.write(f"Brand: {item['brand']}  ·  Model: {item['model']}  ·  Colour code: {item['colour_code']}")
                    st.write(f"Type: {item['type']}  ·  {len(item['barcodes'])} barcode(s) share this")
                    st.divider()
                    cur = assign.get(key, {})
                    for field in item["missing"]:
                        chosen = cur.get(field)
                        st.write(f"**{label_map[field]}**" + (f" → ✅ {chosen}" if chosen else " → _not set_"))
                        palette = palette_map[field]
                        ncols = 6
                        for r in range(0, len(palette), ncols):
                            bcols = st.columns(ncols)
                            for j, colour in enumerate(palette[r:r + ncols]):
                                if bcols[j].button(colour, key=f"cf_{idx}_{field}_{colour}",
                                                   type=("primary" if chosen == colour else "secondary")):
                                    assign.setdefault(key, {})[field] = colour
                                    if all(f in assign.get(key, {}) for f in item["missing"]):
                                        st.session_state["colfill_idx"] = idx + 1
                                    st.rerun()
                    if item.get("can_gradient"):
                        grad_now = bool(assign.get(key, {}).get("__gradient__"))
                        if st.checkbox("Gradient lens", value=grad_now, key=f"cf_grad_{idx}"):
                            assign.setdefault(key, {})["__gradient__"] = True
                        elif key in assign and "__gradient__" in assign[key]:
                            del assign[key]["__gradient__"]

                nav1, nav2, nav3 = st.columns(3)
                if nav1.button("⬅️ Previous", key="cf_prev", disabled=idx == 0):
                    st.session_state["colfill_idx"] = max(0, idx - 1)
                    st.rerun()
                if nav2.button("⏭️ Skip", key="cf_skip"):
                    st.session_state["colfill_idx"] = idx + 1
                    st.rerun()
                if nav3.button("➡️ Next", key="cf_next"):
                    st.session_state["colfill_idx"] = min(total, idx + 1)
                    st.rerun()

        # --- Save ---
        st.divider()
        if assign and st.button(f"💾 Save {len(assign)} assignment(s) to database", type="primary", key="cf_save"):
            with st.spinner("Writing colours to master_catalog..."):
                full = pd.read_sql_table("master_catalog", con=engine)
                full["join_key"] = full["join_key"].astype(str).str.strip()
                bc_by_key = {w["key"]: w["barcodes"] for w in worklist}
                cells = 0
                def _add_gradient(v):
                    parts = [p.strip() for p in str(v or "").split("|") if p.strip() and p.strip().lower() != "nan"]
                    if "Gradient" not in parts:
                        parts.append("Gradient")
                    return "|".join(sorted(set(parts)))

                for gkey, fields in assign.items():
                    barcodes = set(bc_by_key.get(gkey, []))
                    if not barcodes:
                        continue
                    mask = full["join_key"].isin(barcodes)
                    for field, value in fields.items():
                        if field == "__gradient__":
                            # Append 'Gradient' to Glasses_lens_effect, keeping existing effects
                            if "Glasses_lens_effect" not in full.columns:
                                full["Glasses_lens_effect"] = ""
                            full.loc[mask, "Glasses_lens_effect"] = full.loc[mask, "Glasses_lens_effect"].apply(_add_gradient)
                            cells += int(mask.sum())
                            continue
                        if field not in full.columns:
                            full[field] = ""
                        full.loc[mask, field] = value
                        cells += int(mask.sum())
                full.to_sql("master_catalog", con=engine, if_exists="replace", index=False)
                st.success(f"✅ Saved. Updated {cells} cell(s) across {len(assign)} model/colour group(s).")
                for k in ("colfill_worklist", "colfill_idx", "colfill_assign", "colfill_zip_bytes"):
                    st.session_state.pop(k, None)
                st.rerun()

# ==========================================
# ✏️ RENAME PRODUCTS BY BARCODE
# ==========================================
st.divider()
st.subheader("✏️ Rename Products by Barcode")
st.caption(
    "Upload a file with barcodes and names. For every barcode found in the "
    "database, the product name is updated. Great for bulk-fixing names. "
    "(Reuses the same barcode normalization, so leading zeros / .0 don't matter.)"
)

with st.expander("✏️ Open the renamer", expanded=False):
    rename_file = st.file_uploader("Upload file with Barcode + Name columns", type=["xlsx", "csv"], key="renamer_file")
    if rename_file:
        try:
            rdf = _read_any(rename_file)
        except Exception as e:
            st.error(f"Could not read file: {e}")
            rdf = None

        if rdf is not None:
            bc_col = _find_col(rdf, ["Barcode", "EAN", "UPC", "EAN/UPC"])
            name_col = _find_col(rdf, ["Glasses name", "Name", "Assembled_Name", "XML description", "Product name"])

            c1, c2 = st.columns(2)
            bc_col = c1.selectbox("Barcode column", list(rdf.columns),
                                  index=(list(rdf.columns).index(bc_col) if bc_col else 0), key="renamer_bc")
            name_col = c2.selectbox("Name column", list(rdf.columns),
                                    index=(list(rdf.columns).index(name_col) if name_col else 0), key="renamer_name")

            preview = pd.DataFrame({
                "Barcode": rdf[bc_col].astype(str).str.strip(),
                "New name": rdf[name_col].astype(str).str.strip(),
            })
            preview = preview[(preview["Barcode"] != "") & (preview["Barcode"].str.lower() != "nan")]
            st.caption(f"{len(preview)} rows to apply. Preview:")
            st.dataframe(preview.head(10), use_container_width=True, hide_index=True)

            if st.button("✏️ Apply renames to database", type="primary", key="renamer_apply"):
                with st.spinner("Updating names..."):
                    full = pd.read_sql_table("master_catalog", con=engine)
                    full["join_key"] = full["join_key"].astype(str).str.strip()
                    if "Assembled_Name" not in full.columns:
                        full["Assembled_Name"] = ""
                    # Build barcode -> new name map (last wins on dupes)
                    name_map = {}
                    for _, r in preview.iterrows():
                        key = _clean_bc(r["Barcode"])
                        nm = str(r["New name"]).strip()
                        if key and key != "nan" and nm and nm.lower() != "nan":
                            name_map[key] = nm

                    mask = full["join_key"].isin(name_map.keys())
                    updated = 0
                    for pos in full.index[mask]:
                        key = full.at[pos, "join_key"]
                        full.at[pos, "Assembled_Name"] = name_map[key]
                        updated += 1

                    found_keys = set(full.loc[mask, "join_key"])
                    not_found = [bc for bc, key in
                                 ((str(r["Barcode"]).strip(), _clean_bc(r["Barcode"])) for _, r in preview.iterrows())
                                 if key not in found_keys]
                    not_found = sorted(set(not_found))

                    full.to_sql("master_catalog", con=engine, if_exists="replace", index=False)
                    st.success(f"✅ Renamed {updated} product row(s) across {len(name_map)} barcode(s).")
                    if not_found:
                        st.warning(f"⚠️ {len(not_found)} barcode(s) were not found in the database.")
                        st.download_button(
                            "📥 Download not-found barcodes (CSV)",
                            data=pd.DataFrame({"Barcode": not_found}).to_csv(index=False).encode("utf-8-sig"),
                            file_name="rename_not_found.csv",
                            mime="text/csv",
                        )