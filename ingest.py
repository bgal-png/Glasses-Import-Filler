# -*- coding: utf-8 -*-
"""
Pure (Streamlit-free) ingest logic.

Used by:
  - app_admin.py (wraps with Streamlit UI for interactive use)
  - scripts/auto_ingest_safilo.py (headless daily automation via GitHub Actions)

Both call the SAME code paths so behavior stays in sync.
"""
import pandas as pd
import re
from dictionaries import (
    MANUFACTURER_CONFIG,
    VALUE_TRANSLATOR,
    KNOWN_BRANDS,
    classify_color,
)


# ==========================================================================
# MARCOLIN — NEW MASTER FORMAT (June 2026 onward)
# ==========================================================================
# Completely different layout from the old Marcolin file. Self-contained
# row-by-row translator returning the same (df, unmapped, skipped) shape as
# load_single_catalog so it plugs into the admin app and auto-ingest unchanged.

_MARCOLIN_BRAND_MAP = {
    "guess": "Guess",
    "guess jeans": "Guess",
    "guess by marciano": "Guess",
    "max &co": "Max&Co.",
    "max&co": "Max&Co.",
    "maxmara": "Max Mara",
    "max mara": "Max Mara",
    "adidas sport": "Adidas",
    "adidas originals": "Adidas",
    "adidas": "Adidas",
    "tom ford": "Tom Ford",
    "moncler": "Moncler",
}

_MARCOLIN_SHAPE_MAP = {
    "SQUARE": "Square",
    "RECTANGULAR": "Rectangular",
    "ROUND": "Round",
    "GEOMETRIC": "Extravagant",
    "CAT": "Cat Eye",
    "NAVIGATOR": "Pilot",
    "PILOT": "Pilot",
    "SHIELD": "Single lens",
    "OVAL": "Oval / Elipse",
    "BUTTERFLY": "Butterfly",
    "BROWLINE": "Browline",
    "LECTOR / READING SPECTACLES": "Panthos / Tea cup",
    "LECTOR / READING SPECTACL": "Panthos / Tea cup",
}

_MARCOLIN_RIM_MAP = {
    "FULL RIM": "Full rim",
    "SEMIRIMLESS": "Half rim",
    "RIMLESS": "Rimless",
    "THREE PIECES WITH SCREWS": "Rimless",
    "COMPRESSION THREE PIECES": "Rimless",
    "SHIELD": "Full rim",
}

_MARCOLIN_MATERIAL_MAP = {
    "ACETATE": "Plastic",
    "INJECTED": "Plastic",
    "METAL": "Metal",
    "MAGNESIUM": "Metal",
    "TITANIUM": "Titanium",
    "ALUMINUM": "Metal",
    "NYLON": "Plastic",
}

_MARCOLIN_LENS_MATERIAL_MAP = {
    "POLICARBON": "Polycarbonate",
    "CR39": "CR 39",
    "NYLON": "Nylon",
    "TRIACETATO": "Plastic",
    "NXT": "Plastic",
}

_MARCOLIN_ORIGIN_MAP = {
    "CN": "China",
    "VN": "Vietnam",
    "BD": "Bangladesh",
    "IT": "Italy",
    "JP": "Japan",
    "FR": "France",
}

# F/M/U are certain. The rest (X/G/Z/B/Y/K) are Marcolin/Guess fashion-attribute
# codes whose legend we don't have yet — left empty and flagged so they surface
# in the validator rather than being guessed wrong.
_MARCOLIN_GENDER_MAP = {
    "F": "Woman",
    "M": "Man",
    "U": "Man|Woman",
}


def _marcolin_round(v):
    s = str(v if v is not None else "").strip()
    if not s or s.lower() == "nan":
        return ""
    try:
        clean = re.sub(r"[^\d,.-]", "", s).replace(",", ".")
        return str(int(round(float(clean)))) if clean else s
    except Exception:
        return s


def _marcolin_filter_category(raw):
    """Parse LENS FILTER CATEGORIES -> (category_string, is_polarized).
    Handles '3', '3P', '1/3', 'S3', 'S1/S4', '' etc."""
    s = str(raw or "").strip().upper()
    if not s or s == "NAN":
        return "", False
    polarized = "P" in s
    s = s.replace("P", "").replace("S", "").strip()
    if not s:
        return "", polarized
    if "/" in s:
        parts = [p.strip() for p in s.split("/") if p.strip().isdigit()]
        if len(parts) == 2:
            lo, hi = sorted(parts, key=int)
            return f"Category range {lo} - {hi}", polarized
        if len(parts) == 1:
            return f"Category {parts[0]}", polarized
        return "", polarized
    if s.isdigit():
        return f"Category {s}", polarized
    return "", polarized


def _load_marcolin_new(df):
    unmapped = set()
    skipped = set()
    rows = []

    for _, src in df.iterrows():
        barcode = str(src.get("UPC", "")).strip()
        if not barcode or barcode.lower() == "nan":
            continue
        join_key = re.sub(r"\.0$", "", barcode).lstrip("0")
        if not join_key or join_key == "nan":
            continue

        out = {"Barcode": barcode, "join_key": join_key}

        # ---- Brand / Manufacturer ----
        brand_raw = str(src.get("Brand.1", "")).strip()
        brand_norm = re.sub(r"\s+", " ", brand_raw).strip().lower()
        brand = _MARCOLIN_BRAND_MAP.get(brand_norm, brand_raw)
        if brand_norm and brand_norm not in _MARCOLIN_BRAND_MAP:
            unmapped.add(f"Marcolin -> Brand: '{brand_raw}'")
        out["Brand"] = brand
        out["Manufacturer"] = brand

        # ---- MAIN MATERIAL encodes type + material + clip-on ----
        main_mat = str(src.get("MAIN MATERIAL", "")).strip().upper()
        if "SUNGLASS" in main_mat or "MASK" in main_mat:
            g_type = "Sunglasses"
        elif "FRAME" in main_mat:
            g_type = "Frames"
        else:
            g_type = ""
        out["Glasses_type"] = g_type

        # ---- Material (prefer FRONT MATERIAL, fall back to MAIN MATERIAL word) ----
        front_mat = str(src.get("FRONT MATERIAL", "")).strip().upper()
        mat_key = front_mat.split("/")[0].strip()  # "INJECTED / METAL" -> "INJECTED"
        material = _MARCOLIN_MATERIAL_MAP.get(mat_key, "")
        if not material:
            for word, mapped in _MARCOLIN_MATERIAL_MAP.items():
                if word in main_mat:
                    material = mapped
                    break
        if not material and mat_key and mat_key not in ("", "NO FRONT", "NAN"):
            unmapped.add(f"Marcolin -> Glasses_main_material: '{front_mat}'")
        out["Glasses_main_material"] = material

        # ---- Dimensions ----
        size = _marcolin_round(src.get("SIZE"))
        out["Glasses_size_lens_width"] = size
        out["Combination"] = size
        out["Glasses_size_bridge"] = _marcolin_round(src.get("DBL"))
        out["Glasses_size_temple_length"] = _marcolin_round(src.get("TEMPLE"))
        out["Glasses_size_lens_height"] = _marcolin_round(src.get("B MEASURE"))

        # ---- Shape ----
        shape = str(src.get("SHAPE", "")).strip().upper()
        out["Glasses_shape"] = _MARCOLIN_SHAPE_MAP.get(shape, "")
        if shape and shape != "NAN" and shape not in _MARCOLIN_SHAPE_MAP:
            unmapped.add(f"Marcolin -> Glasses_shape: '{shape}'")

        # ---- Rim ----
        rim = str(src.get("TYPOLOGY", "")).strip().upper()
        out["Glasses_frame_type"] = _MARCOLIN_RIM_MAP.get(rim, "")
        if rim and rim != "NAN" and rim not in _MARCOLIN_RIM_MAP:
            unmapped.add(f"Marcolin -> Glasses_frame_type: '{rim}'")

        # ---- Flex ----
        flex = str(src.get("FLEX", "")).strip().upper()
        out["Glasses_other_info"] = "Flex" if flex == "SI" else ""

        # ---- Gender ----
        g = str(src.get("GENDER", "")).strip().upper()
        out["Glasses_gendre"] = _MARCOLIN_GENDER_MAP.get(g, "")
        if g and g not in _MARCOLIN_GENDER_MAP:
            unmapped.add(f"Marcolin -> Glasses_gendre: '{g}' (needs legend)")

        # ---- Colours (separate columns; classify each) ----
        front_col = str(src.get("FRONT COLOUR", "")).strip()
        temple_col = str(src.get("TEMPLE COLOUR", "")).strip()
        lens_col = str(src.get("LENS COLOR", "")).strip()
        if front_col and front_col.lower() != "nan":
            res = classify_color(front_col, "frame")
            out["Frame_Colour"] = res
            if not res:
                unmapped.add(f"Marcolin -> Frame_Colour: '{front_col}'")
        else:
            out["Frame_Colour"] = ""
        if temple_col and temple_col.lower() != "nan":
            res = classify_color(temple_col, "frame")
            out["Temple_Colour"] = res
            if not res:
                unmapped.add(f"Marcolin -> Temple_Colour: '{temple_col}'")
        else:
            out["Temple_Colour"] = ""
        if lens_col and lens_col.lower() != "nan":
            res = classify_color(lens_col, "lens")
            out["Glasses_lens_Colour"] = res
            if not res:
                unmapped.add(f"Marcolin -> Glasses_lens_Colour: '{lens_col}'")
        else:
            out["Glasses_lens_Colour"] = ""

        # ---- Lens material ----
        lm = str(src.get("LENS MATERIAL", "")).strip().upper()
        out["Glasses_lens_material"] = _MARCOLIN_LENS_MATERIAL_MAP.get(lm, "")
        if lm and lm not in _MARCOLIN_LENS_MATERIAL_MAP and lm != "NAN":
            unmapped.add(f"Marcolin -> Glasses_lens_material: '{lm}'")

        # ---- Filter category (+ polarized signal) ----
        filter_cat, pol_from_filter = _marcolin_filter_category(src.get("LENS FILTER CATEGORIES"))
        out["Sunglasses_filter"] = filter_cat

        # ---- Lens effect ----
        eff = set()
        lens_type = str(src.get("LENSES TYPE DESCRIPTION", "")).strip().upper()
        if "POLAR" in lens_type or pol_from_filter:
            eff.add("Polarized")
        if "PHOTO" in lens_type:
            eff.add("Photochromic")
        if str(src.get("GRADIENT", "")).strip().upper() == "YES":
            eff.add("Gradient")
        if str(src.get("MIRROR COATING", "")).strip().upper() == "YES":
            eff.add("Mirror")
        out["Glasses_lens_effect"] = "|".join(sorted(eff))

        # ---- RX ----
        rx = str(src.get("RX ABILITY", "")).strip().upper()
        out["SunGlasses_RX_lenses"] = "Yes" if rx == "S" else ""

        # ---- Origin ----
        origin = str(src.get("COUTRY OF ORIGIN", "")).strip().upper()
        out["Item_origin_country"] = _MARCOLIN_ORIGIN_MAP.get(origin, origin if origin and origin != "NAN" else "")

        # ---- Weight (kg -> g) ----
        nw = str(src.get("NET WEIGHT", "")).strip()
        if nw and nw.lower() != "nan":
            try:
                out["Glasses_weight_g"] = str(round(float(nw) * 1000))
            except Exception:
                pass

        # ---- Model + colour code (SKU = STYLE@SIZE+COLOR#) ----
        style = str(src.get("STYLE", "")).strip()
        sku = str(src.get("SKU", "")).strip()
        color_code = ""
        m = re.search(r"@(.+?)#?$", sku)
        if m:
            inner = m.group(1)
            color_code = inner[len(size):] if size and inner.startswith(size) else inner
        out["Extracted_Model"] = style
        out["Extracted_Color"] = color_code
        out["Glasses_color_code"] = color_code

        # ---- Clip-on ----
        clip = ""
        if "CLIP-ON" in main_mat or str(src.get("CLIP-ON", "")).strip() == "ClipOn Included":
            clip = "Sun clip-on"
        out["Extracted_Clip_on"] = clip
        out["Clip_on_Alert"] = bool(clip and "Polarized" in out["Glasses_lens_effect"])

        # ---- Assembled name ----
        name_parts = [p for p in (brand, style, color_code) if p and p.lower() != "nan"]
        out["Assembled_Name"] = " ".join(name_parts)

        out["Producing_company"] = "Marcolin"
        rows.append(out)

    result_df = pd.DataFrame(rows)
    return result_df, unmapped, skipped


def load_single_catalog(mfg_name, config_settings, file_path):
    """Apply the manufacturer rules engine to a single raw catalog file.

    Returns a 3-tuple:
        processed_df         — cleaned DataFrame with a `join_key` column ready for upsert
        unmapped_values      — set of "<Mfg> -> <Column>: '<raw>'" strings
        skipped_not_mapped   — set of "<Mfg> -> <Column>" strings whose VALUE_TRANSLATOR
                               entry was the sentinel "NOT MAPPED"

    Raises:
        ValueError if the file cannot be parsed or the Barcode column is missing.
    """
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

        # De-duplicate column names by suffixing repeats with .1, .2, ...
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
        raise ValueError(f"Failed to read file '{file_path}': {e}")

    # Marcolin switched to a new master-file format (June 2026) with completely
    # different column names. Detect it by signature columns and route to the
    # dedicated handler. Old-format Marcolin files (if any) fall through.
    if mfg_name == "marcolin" and {"MAIN MATERIAL", "STYLE", "TYPOLOGY"}.issubset(set(df.columns)):
        return _load_marcolin_new(df)

    new_df = pd.DataFrame()

    # 1. Map Columns
    for global_name, mfg_names in config_settings["columns"].items():
        if not mfg_names:
            continue
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
                def merge_row(row, cols=existing_cols):
                    vals = [
                        str(row[c]).strip()
                        for c in cols
                        if pd.notna(row[c]) and str(row[c]).strip().lower() not in ("nan", "")
                    ]
                    return "|".join(vals) if vals else ""
                new_df[global_name] = df.apply(merge_row, axis=1)

    # 1b. Safilo: Combination is just the lens width (e.g. "56"), not the full
    # lens-bridge-temple string. ItemSize already holds the rounded width.
    if mfg_name == "safilo" and "ItemSize" in df.columns:
        def _build_safilo_combination(row):
            val = str(row.get("ItemSize", "")).strip()
            if not val or val.lower() in ("nan", ""):
                return ""
            try:
                clean = re.sub(r"[^\d,.-]", "", val).replace(",", ".")
                if clean:
                    return str(int(round(float(clean))))
            except Exception:
                pass
            return val
        new_df["Combination"] = df.apply(_build_safilo_combination, axis=1)

    # 1c. Safilo: color code is "<ItemColor>/<LenColor>" — frame color code
    # over lens color code (e.g. "HBN/9K"). If only one is present, emit just
    # that one (no trailing slash). Overrides the simple ItemColor mapping.
    if mfg_name == "safilo" and ("ItemColor" in df.columns or "LenColor" in df.columns):
        def _build_safilo_color_code(row):
            frame = str(row.get("ItemColor", "")).strip()
            lens = str(row.get("LenColor", "")).strip()
            if frame.lower() in ("nan", ""): frame = ""
            if lens.lower() in ("nan", ""): lens = ""
            if frame and lens:
                return f"{frame}/{lens}"
            return frame or lens
        new_df["Glasses_color_code"] = df.apply(_build_safilo_color_code, axis=1)

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
                if pol in ("X", "Y"):
                    clip_val = "Magnetic sun clip-on p"
                else:
                    clip_val = "Magnetic sun clip-on"

        elif mfg_name == "luxottica":
            desc = ""
            for col_name in raw_row.index:
                if "popis" in col_name.lower() and "model" in col_name.lower():
                    desc = str(raw_row[col_name]).strip().upper()
                    break
            if desc == "CLIP ON":
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
                if "magnetic clip-on" in combined_text or "magnetic clip on" in combined_text:
                    clip_val = "Magnetic sun clip-on"
                elif "clip-on" in combined_text or "clip on" in combined_text:
                    if "magnetic" not in combined_text:
                        clip_val = "Sun clip-on"
                if clip_val:
                    if "polarized" in combined_text or pol_lens == "X":
                        alert = True

        extracted_clip_ons.append(clip_val)
        clip_on_alerts.append(alert)

    new_df["Extracted_Clip_on"] = extracted_clip_ons
    new_df["Clip_on_Alert"] = clip_on_alerts

    # 2b. Luxottica Clip-On → Base Model Matching
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
                hinge = str(row.get("Hinge_raw", "")).strip().upper()
                if hinge and hinge not in ("NAN", "NO FLEX", ""):
                    if "FLEX" in hinge:
                        final_values.add("Flex")
                elif pd.notna(row.get("Glasses_model")) and "FLEX" in str(row["Glasses_model"]).upper():
                    final_values.add("Flex")
                shape_raw = str(raw_val).strip().upper()
                if "DOUBLE BRIDGE" in shape_raw:
                    final_values.add("Double bridge")
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

        elif col_name == "Glasses_lens_effect":
            if mfg == "safilo":
                pol = str(row.get("Polarized_raw", "")).strip().upper()
                if pol in ("X", "Y"):
                    final_values.add("Polarized")
                phot = str(row.get("Photochromic_raw", "")).strip().upper()
                if phot in ("X", "Y"):
                    final_values.add("Photochromic")
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
                            if m_val:
                                final_values.add(m_val)
                            matched = True
                    if not matched:
                        unmapped_values.add(f"Luxottica -> {col_name} (Keyword Search): '{raw_eff}'")
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
                            if l_dict[p.lower()]:
                                final_values.add(l_dict[p.lower()])
                        else:
                            unmapped_values.add(f"{mfg.title()} -> {col_name}: '{p}'")

        elif col_name == "SunGlasses_RX_lenses":
            raw_rx = str(row.get(col_name, "")).strip().upper()
            if mfg in ["safilo", "kering", "marcolin"]:
                if raw_rx in ("X", "Y"):
                    final_values.add("Yes")
            elif mfg == "luxottica":
                if raw_rx == "YES":
                    final_values.add("Yes")

        elif col_name == "Glasses_shape" and mfg in ["kering", "marcolin"]:
            raw_shape = str(row.get(col_name, "")).strip()
            if raw_shape and raw_shape.lower() != "nan":
                first_shape = raw_shape.split("/")[0].strip()
                if col_name in VALUE_TRANSLATOR:
                    translation_dict = VALUE_TRANSLATOR[col_name]
                    lower_dict = {str(k).lower(): v for k, v in translation_dict.items() if k}
                    shape_lower = first_shape.lower()
                    if shape_lower in lower_dict:
                        if lower_dict[shape_lower]:
                            final_values.add(lower_dict[shape_lower])
                    else:
                        unmapped_values.add(f"{mfg.title()} -> {col_name}: '{first_shape}'")
                else:
                    final_values.add(first_shape)

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
                        if lower_dict[part_lower]:
                            final_values.add(lower_dict[part_lower])
                    else:
                        unmapped_values.add(f"{mfg.title()} -> {col_name}: '{part}'")
            else:
                final_values.add(raw_val)

        if "NOT MAPPED" in final_values:
            skipped_not_mapped.add(f"{mfg.title()} -> {col_name}")
            final_values.discard("NOT MAPPED")
        return "|".join(sorted(final_values))

    for target_col in new_df.columns:
        if target_col in VALUE_TRANSLATOR or target_col in [
            "Glasses_other_info", "Glasses_lens_effect", "SunGlasses_RX_lenses",
            "Glasses_type", "Glasses_shape", "Sunglasses_filter",
            "Glasses_lens_Colour", "Frame_Colour", "Temple_Colour",
        ]:
            new_df[target_col] = new_df.apply(lambda row: process_cell_strict(row, target_col, mfg_name), axis=1)

    # PRE-CLEAN: EXTRACT "KIDS" FROM BRAND/MFG
    if "Brand" not in new_df.columns:
        new_df["Brand"] = ""
    if "Manufacturer" not in new_df.columns:
        new_df["Manufacturer"] = ""

    new_df["Is_Kids"] = (
        new_df["Brand"].astype(str).str.contains(r"(?i)\bkids\b", regex=True, na=False)
        | new_df["Manufacturer"].astype(str).str.contains(r"(?i)\bkids\b", regex=True, na=False)
    )

    new_df["Brand"] = (
        new_df["Brand"].astype(str)
        .str.replace(r"(?i)\bkids\b", "", regex=True)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )
    new_df["Brand"] = new_df["Brand"].apply(lambda x: str(x).title() if x and x.lower() != "nan" else "")

    new_df["Manufacturer"] = (
        new_df["Manufacturer"].astype(str)
        .str.replace(r"(?i)\bkids\b", "", regex=True)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )
    new_df["Manufacturer"] = new_df["Manufacturer"].apply(lambda x: str(x).title() if x and x.lower() != "nan" else "")

    # WHITELIST BRAND CLEANING
    _brand_lookup = sorted(KNOWN_BRANDS, key=len, reverse=True)

    def _clean_brand_to_whitelist(raw):
        raw = str(raw).strip()
        if not raw or raw.lower() == "nan":
            return raw
        raw_lower = raw.lower()
        for known in _brand_lookup:
            if raw_lower == known.lower():
                return known
            if raw_lower.startswith(known.lower() + " "):
                return known
        return raw

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

    # ASSEMBLE MODEL AND NAMES
    def assemble_name_and_parts(row, mfg):
        brand = str(row.get("Brand", "")).strip()
        is_kids = row.get("Is_Kids", False)
        model_out, color_out = "", ""

        if mfg == "safilo":
            model_out = str(row.get("Glasses_model", "")).strip()
            color_out = str(row.get("Glasses_color_code", "")).strip()
            if model_out.lower() == "nan":
                model_out = ""
            if color_out.lower() == "nan":
                color_out = ""
            if is_kids and model_out:
                model_out = f"Kids {model_out}"
            elif is_kids:
                model_out = "Kids"
            parts = [brand, model_out, color_out]

        elif mfg == "luxottica":
            model_out = str(row.get("Glasses_model", "")).strip().lstrip("0")
            color_out = str(row.get("Glasses_color_code", "")).strip()
            if model_out.lower() == "nan":
                model_out = ""
            if color_out.lower() == "nan":
                color_out = ""
            if is_kids and model_out:
                model_out = f"Kids {model_out}"
            elif is_kids:
                model_out = "Kids"
            parts = [brand, model_out, color_out]

        elif mfg in ["kering", "marcolin"]:
            mat_num = str(row.get("Material_Number", "")).strip()
            if mat_num and mat_num.lower() != "nan":
                first_part = mat_num.split(" ")[0]
                model_color = first_part.replace("-", " ")
                mc_parts = model_color.split(" ")
                model_out = mc_parts[0]
                if is_kids and model_out:
                    model_out = f"Kids {model_out}"
                elif is_kids:
                    model_out = "Kids"
                if len(mc_parts) > 1:
                    color_out = mc_parts[1]
                    parts = [brand, model_out, color_out]
                else:
                    parts = [brand, model_out]
            else:
                if is_kids:
                    model_out = "Kids"
                parts = [brand, model_out] if model_out else [brand]
        else:
            if is_kids:
                model_out = "Kids"
            parts = [brand, model_out] if model_out else [brand]

        final_name = " ".join([p for p in parts if p])
        return final_name, model_out, color_out

    if not new_df.empty:
        temp_col = new_df.apply(lambda row: assemble_name_and_parts(row, mfg_name), axis=1)
        new_df["Assembled_Name"] = temp_col.apply(lambda x: x[0] if isinstance(x, (list, tuple)) else "")
        new_df["Extracted_Model"] = temp_col.apply(lambda x: x[1] if isinstance(x, (list, tuple)) else "")
        new_df["Extracted_Color"] = temp_col.apply(lambda x: x[2] if isinstance(x, (list, tuple)) else "")

    # Hugo Boss sub-brand resolution — Safilo (and any other mfg) lists the
    # group's two sub-lines under different BrandD values. We want customer-
    # facing Brand/Manufacturer to follow the "<line> by Hugo Boss" convention:
    #   Brand "Hugo Boss" + model "BOSS …"  →  "Boss by Hugo Boss"
    #   Brand "Hugo Boss" + model "HG …"    →  "Hugo by Hugo Boss"  (rare edge case)
    #   Brand "Hugo" (standalone)           →  "Hugo by Hugo Boss"
    # Runs AFTER assemble_name_and_parts so the Glasses name keeps the shorter
    # form ("Hugo Boss BOSS 1594 001/9O") while Brand/Manufacturer get long form.
    if not new_df.empty and "Brand" in new_df.columns:
        def _resolve_hugo_boss(row, col):
            val = str(row.get(col, "")).strip()
            model = str(row.get("Glasses_model", "")).strip().upper()
            if val == "Hugo Boss":
                if model.startswith("HG"):
                    return "Hugo by Hugo Boss"
                return "Boss by Hugo Boss"
            if val == "Hugo":
                return "Hugo by Hugo Boss"
            return val
        new_df["Brand"] = new_df.apply(lambda r: _resolve_hugo_boss(r, "Brand"), axis=1)
        new_df["Manufacturer"] = new_df.apply(lambda r: _resolve_hugo_boss(r, "Manufacturer"), axis=1)

    if "Is_Kids" in new_df.columns:
        new_df.drop(columns=["Is_Kids"], inplace=True)

    for dim_col in [
        "Glasses_size_temple_length", "Glasses_size_lens_height",
        "Glasses_size_lens_width", "Glasses_size_bridge",
    ]:
        if dim_col in new_df.columns:
            def round_dimension(val):
                if pd.isna(val) or str(val).strip() == "" or str(val).strip().lower() == "nan":
                    return ""
                try:
                    clean_str = re.sub(r"[^\d,.-]", "", str(val).strip()).replace(",", ".")
                    if clean_str:
                        return str(int(round(float(clean_str))))
                except Exception:
                    pass
                return str(val).strip()
            new_df[dim_col] = new_df[dim_col].apply(round_dimension)

    if "Barcode" in new_df.columns:
        new_df["join_key"] = (
            new_df["Barcode"].astype(str)
            .str.strip()
            .str.replace(r"\.0$", "", regex=True)
            .str.lstrip("0")
        )
        new_df = new_df[new_df["join_key"].notna() & (new_df["join_key"] != "nan") & (new_df["join_key"] != "")]
    else:
        raise ValueError(f"'Barcode' column missing in {mfg_name} after extraction.")

    new_df["Producing_company"] = mfg_name.title()

    return new_df, unmapped_values, skipped_not_mapped


def perform_upsert(new_data_df, engine):
    """Merge a freshly processed DataFrame into the `master_catalog` table.

    - Deduplicates by join_key (keep last).
    - Updates rows that already exist; appends rows that don't.
    - Writes back the union via to_sql replace (full table rewrite — the existing pattern).

    Returns a short status string.
    """
    new_data_df.drop_duplicates(subset=["join_key"], keep="last", inplace=True)
    new_data_df.set_index("join_key", inplace=True)

    try:
        existing_df = pd.read_sql_table("master_catalog", con=engine)
        existing_df.set_index("join_key", inplace=True)

        common_indices = new_data_df.index.intersection(existing_df.index)
        updated_count = len(common_indices)

        existing_df.update(new_data_df)

        new_rows = new_data_df[~new_data_df.index.isin(existing_df.index)]
        added_count = len(new_rows)

        final_df = pd.concat([existing_df, new_rows])

        msg = f"Refreshed {updated_count} existing products. Added {added_count} new products."

    except Exception:
        final_df = new_data_df
        msg = f"Created database from scratch with {len(new_data_df)} products."

    final_df.reset_index().to_sql("master_catalog", con=engine, if_exists="replace", index=False)
    return msg
