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
    "METACRILATO": "Plastic",
    "POLYCARBONATE": "Polycarbonate",
}

_MARCOLIN_ORIGIN_MAP = {
    "CN": "China",
    "VN": "Vietnam",
    "BD": "Bangladesh",
    "IT": "Italy",
    "JP": "Japan",
    "FR": "France",
    "DE": "Germany",
    "KH": "Cambodia",
    "SI": "Slovenia",
}

# Marcolin/Guess gender codes (legend confirmed):
#   F=Female, M=Male, U=Unisex, X/Z=unspecified (unisex),
#   K=Kids, B=Boys, G=Girls, Y=Youth — all four kids variants -> Child.
_MARCOLIN_GENDER_MAP = {
    "F": "Woman",
    "M": "Man",
    "U": "Man|Woman",
    "X": "Man|Woman",
    "Z": "Man|Woman",
    "K": "Child",
    "B": "Child",
    "G": "Child",
    "Y": "Child",
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
            material = front_mat  # keep original, flag for review
        out["Glasses_main_material"] = material

        # ---- Dimensions ----
        size = _marcolin_round(src.get("SIZE"))
        out["Glasses_size_lens_width"] = size
        out["Combination"] = size
        out["Glasses_size_bridge"] = _marcolin_round(src.get("DBL"))
        out["Glasses_size_temple_length"] = _marcolin_round(src.get("TEMPLE"))
        out["Glasses_size_lens_height"] = _marcolin_round(src.get("B MEASURE"))

        # ---- Shape ----
        shape = str(src.get("SHAPE", "")).strip()
        if not shape or shape.upper() == "NAN":
            out["Glasses_shape"] = ""
        elif shape.upper() in _MARCOLIN_SHAPE_MAP:
            out["Glasses_shape"] = _MARCOLIN_SHAPE_MAP[shape.upper()]
        else:
            out["Glasses_shape"] = shape  # keep original, flag for review
            unmapped.add(f"Marcolin -> Glasses_shape: '{shape}'")

        # ---- Rim ----
        rim = str(src.get("TYPOLOGY", "")).strip()
        if not rim or rim.upper() == "NAN":
            out["Glasses_frame_type"] = ""
        elif rim.upper() in _MARCOLIN_RIM_MAP:
            out["Glasses_frame_type"] = _MARCOLIN_RIM_MAP[rim.upper()]
        else:
            out["Glasses_frame_type"] = rim  # keep original, flag for review
            unmapped.add(f"Marcolin -> Glasses_frame_type: '{rim}'")

        # ---- Flex ----
        flex = str(src.get("FLEX", "")).strip().upper()
        out["Glasses_other_info"] = "Flex" if flex == "SI" else ""

        # ---- Gender ----
        g = str(src.get("GENDER", "")).strip().upper()
        if not g or g == "NAN":
            out["Glasses_gendre"] = ""
        elif g in _MARCOLIN_GENDER_MAP:
            out["Glasses_gendre"] = _MARCOLIN_GENDER_MAP[g]
        else:
            out["Glasses_gendre"] = g  # keep original code, flag for review
            unmapped.add(f"Marcolin -> Glasses_gendre: '{g}'")

        # ---- Colours (separate columns; classify each) ----
        front_col = str(src.get("FRONT COLOUR", "")).strip()
        temple_col = str(src.get("TEMPLE COLOUR", "")).strip()
        lens_col = str(src.get("LENS COLOR", "")).strip()
        # Colours: keep the original value when the classifier can't match it,
        # so the cell isn't empty and the validator can flag it for manual fix.
        if front_col and front_col.lower() != "nan":
            res = classify_color(front_col, "frame")
            out["Frame_Colour"] = res if res else front_col
            if not res:
                unmapped.add(f"Marcolin -> Frame_Colour: '{front_col}'")
        else:
            out["Frame_Colour"] = ""
        if temple_col and temple_col.lower() != "nan":
            res = classify_color(temple_col, "frame")
            out["Temple_Colour"] = res if res else temple_col
            if not res:
                unmapped.add(f"Marcolin -> Temple_Colour: '{temple_col}'")
        else:
            out["Temple_Colour"] = ""
        if lens_col and lens_col.lower() != "nan":
            res = classify_color(lens_col, "lens")
            out["Glasses_lens_Colour"] = res if res else lens_col
            if not res:
                unmapped.add(f"Marcolin -> Glasses_lens_Colour: '{lens_col}'")
        else:
            out["Glasses_lens_Colour"] = ""

        # ---- Lens material ----
        lm = str(src.get("LENS MATERIAL", "")).strip()
        if not lm or lm.upper() == "NAN":
            out["Glasses_lens_material"] = ""
        elif lm.upper() in _MARCOLIN_LENS_MATERIAL_MAP:
            out["Glasses_lens_material"] = _MARCOLIN_LENS_MATERIAL_MAP[lm.upper()]
        else:
            out["Glasses_lens_material"] = lm  # keep original, flag for review
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


# ==========================================================================
# TOM FORD catalogue (Marcolin family, Tom Ford-specific column names)
# ==========================================================================
# Same conventions as the Marcolin master (MAIN MATERIAL encodes type,
# SKU = MODEL@SIZE+COLOR#, SI/NO flex, filter categories, +CLIP-ON) but with
# Tom Ford column names (MODEL, SIZE/COLOR, BRIDGE, MADE IN) and no TYPOLOGY /
# lens-height / RX columns. Uploaded as "marcolin"; Producing_company=Marcolin.

def _load_marcolin_tomford(df):
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

        # ---- Brand (always Tom Ford in this file) ----
        brand = str(src.get("BRAND", "")).strip() or "Tom Ford"
        out["Brand"] = brand
        out["Manufacturer"] = brand

        # ---- Type from MAIN MATERIAL (MASK counts as Sunglasses) ----
        main_mat = str(src.get("MAIN MATERIAL", "")).strip().upper()
        if "SUNGLASS" in main_mat or "MASK" in main_mat:
            out["Glasses_type"] = "Sunglasses"
        elif "FRAME" in main_mat:
            out["Glasses_type"] = "Frames"
        else:
            out["Glasses_type"] = ""

        # ---- Material (FRONT MATERIAL, fall back to MAIN MATERIAL word) ----
        front_mat = str(src.get("FRONT MATERIAL", "")).strip().upper()
        mat_key = front_mat.split("/")[0].strip()
        material = _MARCOLIN_MATERIAL_MAP.get(mat_key, "")
        if not material:
            for word, mapped in _MARCOLIN_MATERIAL_MAP.items():
                if word in main_mat:
                    material = mapped
                    break
        if not material and mat_key and mat_key not in ("", "NO FRONT", "NAN"):
            unmapped.add(f"TomFord -> Glasses_main_material: '{front_mat}'")
            material = front_mat
        out["Glasses_main_material"] = material or None

        # ---- Dimensions (lens width, bridge, temple; no lens height) ----
        size = _marcolin_round(src.get("SIZE"))
        out["Glasses_size_lens_width"] = size or None
        out["Combination"] = size or None
        out["Glasses_size_bridge"] = _marcolin_round(src.get("BRIDGE")) or None
        out["Glasses_size_temple_length"] = _marcolin_round(src.get("TEMPLE")) or None

        # ---- Shape ----
        shape = str(src.get("SHAPE", "")).strip()
        if not shape or shape.upper() == "NAN":
            out["Glasses_shape"] = None
        elif shape.upper() in _MARCOLIN_SHAPE_MAP:
            out["Glasses_shape"] = _MARCOLIN_SHAPE_MAP[shape.upper()]
        else:
            out["Glasses_shape"] = shape
            unmapped.add(f"TomFord -> Glasses_shape: '{shape}'")

        # ---- Flex ----
        flex = str(src.get("FLEX", "")).strip().upper()
        out["Glasses_other_info"] = "Flex" if flex == "SI" else None

        # ---- Gender ----
        g = str(src.get("GENDER", "")).strip().upper()
        if not g or g == "NAN":
            out["Glasses_gendre"] = None
        elif g in _MARCOLIN_GENDER_MAP:
            out["Glasses_gendre"] = _MARCOLIN_GENDER_MAP[g]
        else:
            out["Glasses_gendre"] = g
            unmapped.add(f"TomFord -> Glasses_gendre: '{g}'")

        # ---- Colours ----
        for src_col, out_col, ctype in [
            ("FRONT COLOUR", "Frame_Colour", "frame"),
            ("TEMPLE COLOUR", "Temple_Colour", "frame"),
            ("LENS COLOR", "Glasses_lens_Colour", "lens"),
        ]:
            v = str(src.get(src_col, "")).strip()
            if v and v.lower() != "nan":
                res = classify_color(v, ctype)
                out[out_col] = res if res else v
                if not res:
                    unmapped.add(f"TomFord -> {out_col}: '{v}'")
            else:
                out[out_col] = None

        # ---- Lens material ----
        lm = str(src.get("LENS MATERIAL", "")).strip()
        lm_key = lm.split("/")[0].strip().upper()  # "CR39/NYLON" -> "CR39"
        if not lm or lm.upper() == "NAN":
            out["Glasses_lens_material"] = None
        elif lm_key in _MARCOLIN_LENS_MATERIAL_MAP:
            out["Glasses_lens_material"] = _MARCOLIN_LENS_MATERIAL_MAP[lm_key]
        else:
            out["Glasses_lens_material"] = lm
            unmapped.add(f"TomFord -> Glasses_lens_material: '{lm}'")

        # ---- Filter category + lens effect ----
        filter_cat, pol_from_filter = _marcolin_filter_category(src.get("LENS FILTER CATEGORIES"))
        out["Sunglasses_filter"] = filter_cat or None
        eff = set()
        lens_type = str(src.get("LENSES TYPE DESCRIPTION", "")).strip().upper()
        if "POLAR" in lens_type or pol_from_filter:
            eff.add("Polarized")
        if "PHOTO" in lens_type:
            eff.add("Photochromic")
        if str(src.get("GRADIENT", "")).strip().upper() == "YES":
            eff.add("Gradient")
        out["Glasses_lens_effect"] = "|".join(sorted(eff)) if eff else None

        # ---- Origin (MADE IN) ----
        origin = str(src.get("MADE IN", "")).strip().upper()
        out["Item_origin_country"] = _MARCOLIN_ORIGIN_MAP.get(
            origin, origin if origin and origin != "NAN" else None
        )

        # ---- Model + colour code (SKU = MODEL@SIZE+COLOR#) ----
        model = str(src.get("MODEL", "")).strip()
        size_color = str(src.get("SIZE/COLOR", "")).strip()
        color_code = size_color[len(size):] if size and size_color.startswith(size) else size_color
        out["Extracted_Model"] = model
        out["Extracted_Color"] = color_code
        out["Glasses_color_code"] = color_code

        # ---- Clip-on ----
        clip = ""
        if str(src.get("CLIP-ON", "")).strip() == "ClipOn Included" or "CLIP-ON" in main_mat:
            polarized = "Polarized" in (out["Glasses_lens_effect"] or "")
            clip = "Magnetic sun clip-on p" if polarized else "Magnetic sun clip-on"
        out["Extracted_Clip_on"] = clip
        out["Clip_on_Alert"] = False

        # ---- Name ----
        name_parts = [p for p in (brand, model, color_code) if p and p.lower() != "nan"]
        out["Assembled_Name"] = " ".join(name_parts)

        out["Producing_company"] = "Marcolin"
        rows.append(out)

    return pd.DataFrame(rows), unmapped, skipped


# ==========================================================================
# THÉLIOS — LVMH eyewear master data (Celine, Dior, Bulgari, Loewe, Fendi,
# Barton Perreira, Vuarnet, Givenchy, Tag Heuer, Kenzo, Fred, ...)
# ==========================================================================
# New manufacturer. Well-structured English columns, but the export carries a
# junk top row (real headers on the 2nd row) which the handler promotes.

_THELIOS_BRAND_MAP = {
    "celine": "Celine",
    "dior woman": "Dior",
    "dior man": "Dior",
    "dior": "Dior",
    "bulgari": "Bulgari",
    "loewe": "Loewe",
    "fendi": "Fendi",
    "vuarnet": "Vuarnet",
    "givenchy": "Givenchy",
    "tag heuer": "Tag Heuer",
    "kenzo": "Kenzo",
    "fred": "Fred",
    "barton perreira": "Barton Perreira",
}

_THELIOS_SHAPE_MAP = {
    "GEOMETRIC": "Extravagant",
    "RECTANGULAR": "Rectangular",
    "SQUARE": "Square",
    "ROUND": "Round",
    "OVAL": "Oval / Elipse",
    "CAT EYE": "Cat Eye",
    "PILOT": "Pilot",
    "NAVIGATOR": "Pilot",
    "AVIATOR": "Pilot",
    "BUTTERFLY": "Butterfly",
    "SHIELD": "Single lens",
    "MASK": "Single lens",
    "PANTOS": "Panthos / Tea cup",
    "PANTHOS": "Panthos / Tea cup",
    "BROWLINE": "Browline",
}

_THELIOS_RIM_MAP = {
    "FULL RIM": "Full rim",
    "RIMLESS": "Rimless",
    "SEMIRIMLESS": "Half rim",
    "INVERTED HALF RIM": "Half rim",
    "HALF RIM": "Half rim",
    "3 PIECES COMPRESSION": "Rimless",
    "SHIELD": "Full rim",
}

_THELIOS_GENDER_MAP = {
    "FEMALE": "Woman",
    "MALE": "Man",
    "UNISEX": "Man|Woman",
    "MAN": "Man",
    "WOMAN": "Woman",
    "KIDS": "Child",
    "JUNIOR": "Child",
}


def _thelios_material(raw):
    s = str(raw or "").strip()
    low = s.lower()
    if not s or low == "nan":
        return None
    if "titanium" in low:
        return "Titanium"
    if any(k in low for k in ("acetate", "injected", "nylon", "plastic")):
        return "Plastic"
    if any(k in low for k in ("metal", "alumin", "gold", "steel", "monel")):
        return "Metal"
    return None


def _thelios_filter(raw):
    """Parse Lens Base -> (category, is_polarized). Handles 3, 3P, 3L, 3PL,
    1-2, 1-3, 0..4. L (mirror/light treatment) is ignored for the category."""
    s = str(raw or "").strip().upper()
    if not s or s == "NAN":
        return "", False
    pol = "P" in s
    m = re.search(r"(\d)\s*-\s*(\d)", s)
    if m:
        lo, hi = sorted([m.group(1), m.group(2)], key=int)
        return f"Category range {lo} - {hi}", pol
    m2 = re.search(r"(\d)", s)
    if m2:
        return f"Category {m2.group(1)}", pol
    return "", pol


def _load_thelios(df):
    unmapped = set()
    skipped = set()

    # Promote the real header row if the export carried a junk top row.
    df = df.copy()
    if "Brand" not in df.columns and (df.iloc[0].astype(str) == "Brand").any():
        df.columns = df.iloc[0].astype(str).str.replace("\xa0", " ", regex=False).str.strip()
        df = df.iloc[1:].reset_index(drop=True)
    else:
        df.columns = df.columns.astype(str).str.replace("\xa0", " ", regex=False).str.strip()
    # Drop duplicate columns (junk headers may collide)
    df = df.loc[:, ~df.columns.duplicated()]

    ean_col = next((c for c in df.columns if "EAN" in c.upper()), "* EAN Code")
    rows = []

    for _, src in df.iterrows():
        barcode = str(src.get(ean_col, "")).strip()
        if not barcode or barcode.lower() == "nan":
            continue
        join_key = re.sub(r"\.0$", "", barcode).lstrip("0")
        if not join_key or join_key == "nan":
            continue

        out = {"Barcode": barcode, "join_key": join_key}

        # ---- Brand ----
        brand_raw = str(src.get("Brand", "")).strip()
        brand_norm = re.sub(r"\s+", " ", brand_raw).strip().lower()
        brand = _THELIOS_BRAND_MAP.get(brand_norm, brand_raw)
        if brand_norm and brand_norm not in _THELIOS_BRAND_MAP:
            unmapped.add(f"Thelios -> Brand: '{brand_raw}'")
        out["Brand"] = brand
        out["Manufacturer"] = brand

        # ---- Type + material (from Material: '<mat> Sunglasses/Ophtalmic') ----
        material_raw = str(src.get("Material", "")).strip()
        out["Glasses_type"] = "Sunglasses" if "sunglass" in material_raw.lower() else "Frames"
        is_sun = out["Glasses_type"] == "Sunglasses"
        out["Glasses_main_material"] = _thelios_material(material_raw)

        # ---- Model + colour code (Supplier Color Code = SIZE + COLOR) ----
        model = str(src.get("Model Code", "")).strip()
        supplier_color = str(src.get("Supplier Color Code", "")).strip()
        m = re.match(r"(\d{2})(.*)", supplier_color)
        size_from_code = m.group(1) if m else ""
        color_code = m.group(2).strip() if m else supplier_color
        out["Extracted_Model"] = model
        out["Extracted_Color"] = color_code
        out["Glasses_color_code"] = color_code

        # ---- Dimensions (A=lens width, B=lens height, DBL=bridge, temple) ----
        lens_w = _marcolin_round(src.get("Width (A)")) or size_from_code
        out["Glasses_size_lens_width"] = lens_w or None
        out["Combination"] = lens_w or None
        out["Glasses_size_lens_height"] = _marcolin_round(src.get("Height (B) B-depth")) or None
        out["Glasses_size_bridge"] = _marcolin_round(src.get("Bridge Size - DBL")) or None
        out["Glasses_size_temple_length"] = _marcolin_round(src.get("Temple Length")) or None

        # ---- Shape ----
        shp = str(src.get("Frame Shape", "")).strip()
        if not shp or shp.upper() in ("NAN", "?"):
            out["Glasses_shape"] = None
        elif shp.upper() in _THELIOS_SHAPE_MAP:
            out["Glasses_shape"] = _THELIOS_SHAPE_MAP[shp.upper()]
        else:
            out["Glasses_shape"] = shp
            unmapped.add(f"Thelios -> Glasses_shape: '{shp}'")

        # ---- Rim ----
        rim = str(src.get("Technical Rim Type", "")).strip()
        if not rim or rim.upper() in ("NAN", "?"):
            out["Glasses_frame_type"] = None
        elif rim.upper() in _THELIOS_RIM_MAP:
            out["Glasses_frame_type"] = _THELIOS_RIM_MAP[rim.upper()]
        else:
            out["Glasses_frame_type"] = rim
            unmapped.add(f"Thelios -> Glasses_frame_type: '{rim}'")

        # ---- Gender ----
        g = str(src.get("Gender", "")).strip()
        if not g or g.upper() == "NAN":
            out["Glasses_gendre"] = None
        elif g.upper() in _THELIOS_GENDER_MAP:
            out["Glasses_gendre"] = _THELIOS_GENDER_MAP[g.upper()]
        else:
            out["Glasses_gendre"] = g
            unmapped.add(f"Thelios -> Glasses_gendre: '{g}'")

        # ---- Colour: 'front / lens' (split on ' / ') ----
        color_full = str(src.get("Color", "")).strip()
        front_part, lens_part = color_full, ""
        if color_full and color_full.lower() != "nan":
            parts = re.split(r"\s+/\s+", color_full, maxsplit=1)
            front_part = parts[0].strip()
            lens_part = parts[1].strip() if len(parts) > 1 else ""
        if front_part and front_part.lower() != "nan":
            res = classify_color(front_part, "frame")
            out["Frame_Colour"] = res if res else front_part
            out["Temple_Colour"] = res if res else front_part
            if not res:
                unmapped.add(f"Thelios -> Frame_Colour: '{front_part}'")
        else:
            out["Frame_Colour"] = None
            out["Temple_Colour"] = None
        if is_sun and lens_part:
            res = classify_color(lens_part, "lens")
            out["Glasses_lens_Colour"] = res if res else None
        else:
            out["Glasses_lens_Colour"] = None

        # ---- Filter category + lens effect ----
        filter_cat, pol_lb = _thelios_filter(src.get("Lens Base"))
        out["Sunglasses_filter"] = filter_cat or None
        eff = set()
        if str(src.get("Polarized", "")).strip().lower() == "yes" or pol_lb or "polariz" in lens_part.lower():
            eff.add("Polarized")
        if str(src.get("Photochromic", "")).strip().lower() == "yes":
            eff.add("Photochromic")
        if "gradient" in lens_part.lower():
            eff.add("Gradient")
        if "mirror" in lens_part.lower():
            eff.add("Mirror")
        out["Glasses_lens_effect"] = "|".join(sorted(eff)) if eff else None

        # ---- RX (Glazeable, sunglasses only) ----
        if is_sun and str(src.get("Glazeable", "")).strip().lower() == "yes":
            out["SunGlasses_RX_lenses"] = "Yes"
        else:
            out["SunGlasses_RX_lenses"] = None

        # ---- Origin ----
        origin = str(src.get("Country of Origin", "")).strip().upper()
        out["Item_origin_country"] = _MARCOLIN_ORIGIN_MAP.get(
            origin, origin if origin and origin != "NAN" else None
        )

        # ---- Name ----
        name_parts = [p for p in (brand, model, color_code) if p and p.lower() != "nan"]
        out["Assembled_Name"] = " ".join(name_parts)

        out["Extracted_Clip_on"] = ""
        out["Clip_on_Alert"] = False
        out["Producing_company"] = "Thelios"
        rows.append(out)

    return pd.DataFrame(rows), unmapped, skipped


# ==========================================================================
# SAFILO — optical-frames catalog export (a 2nd Safilo format)
# ==========================================================================
# Distinct from the daily-availability CSV. An optical-frames catalog with
# real Shape data but NO bridge / lens-height / rim columns. Used to top up
# the DB with frames not present in the daily feed. IMPORTANT: this handler
# outputs ONLY the fields the file actually provides, so the barcode upsert
# does not overwrite richer daily-feed data (bridge, lens height, rim, lens
# fields) for overlapping items — those columns are simply absent here.

_SAFILO_FRAMES_BRAND_CORRECTIONS = {
    "moschino love": "Love Moschino",
    "prive' revaux": "Prive Revaux",
    "prive revaux": "Prive Revaux",
    "hugo boss": "Boss by Hugo Boss",
    "hugo": "Hugo by Hugo Boss",
}

_SAFILO_FRAMES_ORIGIN_MAP = {
    "CN": "China", "IT": "Italy", "SI": "Slovenia", "VN": "Vietnam",
    "BD": "Bangladesh", "KH": "Cambodia", "JP": "Japan", "FR": "France",
}


def _safilo_frames_brand(raw):
    s = re.sub(r"(?i)\bkids\b", "", str(raw or "")).strip()
    s = re.sub(r"\s+", " ", s).strip()
    if not s or s.lower() == "nan":
        return ""
    low = s.lower()
    if low in _SAFILO_FRAMES_BRAND_CORRECTIONS:
        return _SAFILO_FRAMES_BRAND_CORRECTIONS[low]
    for known in sorted(KNOWN_BRANDS, key=len, reverse=True):
        if low == known.lower() or low.startswith(known.lower() + " "):
            return known
    return s.title()


def _safilo_frames_material(raw):
    """Keyword-map First Front Material Description to Plastic/Metal/Titanium."""
    s = str(raw or "").strip()
    low = s.lower()
    if not s or low == "nan":
        return None, True
    if "titanium" in low:
        return "Titanium", True
    if any(k in low for k in ("steel", "monel", "metal")):
        return "Metal", True
    if any(k in low for k in ("acetate", "cellulose", "pmma", "polyamide", "polyester",
                              "optyl", "inject", "nylon", "grilamid", "propionate", "prop",
                              "co-polyester", "tr90", "plastic", "polycarbon", "rubber", "tpu",
                              "carbon")):
        return "Plastic", True
    return s, False  # unknown — keep original, flag


def _safilo_lens_material_from_desc(desc):
    """Extract lens material from a Lens Description ('POLYESTER', 'POLICARBONATE
    LENS', 'TRIACETATE LENS'...). Returns None if the description is a colour."""
    d = str(desc or "").upper()
    if "POLICARBON" in d or "POLYCARBON" in d:
        return "Polycarbonate"
    if "TRIACETAT" in d:
        return "Plastic"
    if "POLYESTER" in d:
        return "Plastic"
    if "NYLON" in d:
        return "Nylon"
    if "CR39" in d or "CR 39" in d:
        return "CR 39"
    if "GLASS" in d:
        return "Glass"
    return None


def _safilo_lens_effect_from_desc(desc, photochromic_flag):
    """Derive lens effect set from Lens Description + Photochromic column."""
    d = str(desc or "").upper()
    eff = set()
    if "POLARIZED" in d or "POLARIS" in d:
        eff.add("Polarized")
    if "MIRROR" in d:
        eff.add("Mirror")
    if "SHADED" in d or "GRADIENT" in d:
        eff.add("Gradient")
    if str(photochromic_flag or "").strip().upper() == "X" or "PHOTOCHROM" in d or "PHOTOCROM" in d:
        eff.add("Photochromic")
    return eff


def _load_safilo_catalog(df):
    """Handle the Safilo catalog-export format (optical frames, opt+clip-on, and
    sunglasses — all share the same 27 columns). Type is derived from Product
    type. Empty cells are emitted as None so the barcode upsert SKIPS them,
    preserving richer daily-feed data (bridge/lens height/rim/filter) for
    overlapping items while still adding items missing from the daily feed."""
    unmapped = set()
    skipped = set()
    rows = []

    shape_dict = {str(k).lower(): v for k, v in VALUE_TRANSLATOR.get("Glasses_shape", {}).items() if k}
    type_dict = {str(k).lower(): v for k, v in VALUE_TRANSLATOR.get("Glasses_type", {}).items() if k}

    for _, src in df.iterrows():
        barcode = str(src.get("EAN/UPC", "")).strip()
        if not barcode or barcode.lower() == "nan":
            continue
        join_key = re.sub(r"\.0$", "", barcode).lstrip("0")
        if not join_key or join_key == "nan":
            continue

        out = {"Barcode": barcode, "join_key": join_key}

        # ---- Brand ----
        brand = _safilo_frames_brand(src.get("Division Description"))
        out["Brand"] = brand
        out["Manufacturer"] = brand

        # ---- Type (from Product type) ----
        prod_type = re.sub(r"\s+", " ", str(src.get("Product type", "")).strip().upper())
        g_type = type_dict.get(prod_type.lower(), "")
        if not g_type:
            if "SUNGLASS" in prod_type:
                g_type = "Sunglasses"
            elif "FRAME" in prod_type or "CLIP" in prod_type:
                g_type = "Frames"
        out["Glasses_type"] = g_type
        is_sun = g_type == "Sunglasses"

        # ---- Model + colour code ----
        style = str(src.get("Style", "")).strip()
        color_code = str(src.get("Color Code", "")).strip()
        if color_code.lower() == "nan":
            color_code = ""
        out["Extracted_Model"] = style
        out["Glasses_color_code"] = color_code
        out["Extracted_Color"] = color_code

        # ---- Dimensions we HAVE (lens width + temple only) ----
        size = _marcolin_round(src.get("Size"))
        out["Glasses_size_lens_width"] = size or None
        out["Combination"] = size or None
        out["Glasses_size_temple_length"] = _marcolin_round(src.get("Temple Length")) or None
        # bridge, lens height, rim, filter category intentionally NOT output
        # (absent/unreliable in this format) so the upsert won't wipe daily-feed
        # values for overlapping items.

        # ---- Shape (translate, fall through original) ----
        shp = str(src.get("Shape", "")).strip()
        if not shp or shp.lower() == "nan":
            out["Glasses_shape"] = None
        elif shp.lower() in shape_dict:
            mapped = shape_dict[shp.lower()]
            out["Glasses_shape"] = mapped if mapped and mapped != "NOT MAPPED" else None
        else:
            out["Glasses_shape"] = shp
            unmapped.add(f"Safilo(catalog) -> Glasses_shape: '{shp}'")

        # ---- Frame material ----
        mat, ok = _safilo_frames_material(src.get("First Front Material Description"))
        out["Glasses_main_material"] = mat
        if mat and not ok:
            unmapped.add(f"Safilo(catalog) -> Glasses_main_material: '{src.get('First Front Material Description')}'")

        # ---- Frame colour (front + temple) ----
        col = str(src.get("Color Code Description", "")).strip()
        if col and col.lower() != "nan":
            res = classify_color(col, "frame")
            out["Frame_Colour"] = res if res else col
            out["Temple_Colour"] = res if res else col
            if not res:
                unmapped.add(f"Safilo(catalog) -> Frame_Colour: '{col}'")
        else:
            out["Frame_Colour"] = None
            out["Temple_Colour"] = None

        # ---- Lens fields (sunglasses only; None otherwise so upsert skips) ----
        lens_desc = str(src.get("Lens Description", "")).strip()
        photo = src.get("Photochromic", "")
        if is_sun and lens_desc and lens_desc.lower() != "nan":
            lens_color = classify_color(lens_desc, "lens")
            out["Glasses_lens_Colour"] = lens_color or None  # material-only desc -> None
            out["Glasses_lens_material"] = _safilo_lens_material_from_desc(lens_desc)
            eff = _safilo_lens_effect_from_desc(lens_desc, photo)
            out["Glasses_lens_effect"] = "|".join(sorted(eff)) if eff else None
        else:
            out["Glasses_lens_Colour"] = None
            out["Glasses_lens_material"] = None
            out["Glasses_lens_effect"] = None

        # ---- Origin ----
        origin = str(src.get("Country of origin", "")).strip().upper()
        out["Item_origin_country"] = _SAFILO_FRAMES_ORIGIN_MAP.get(
            origin, origin if origin and origin != "NAN" else None
        )

        # ---- Clip-on (Product type '+ CLIP-ON' or Style ending in /C) ----
        clip = ""
        clip_lens = None
        if "CLIP-ON" in prod_type or "CLIP ON" in prod_type or style.upper().endswith("/C"):
            polarized = "POLARIZED" in lens_desc.upper()
            clip = "Magnetic sun clip-on p" if polarized else "Magnetic sun clip-on"
            clip_lens = classify_color(lens_desc, "lens") or None
        out["Extracted_Clip_on"] = clip
        out["Clip_on_Alert"] = False
        if clip_lens:
            out["Clip_on_lens_colour"] = clip_lens

        # ---- Name ----
        name_parts = [p for p in (brand, style, color_code) if p and p.lower() != "nan"]
        out["Assembled_Name"] = " ".join(name_parts)

        out["Producing_company"] = "Safilo"
        rows.append(out)

    return pd.DataFrame(rows), unmapped, skipped


# ==========================================================================
# DE RIGO — master data format
# ==========================================================================
# New manufacturer (Police, Furla, Just Cavalli). Clean structure: Sun/optical
# flag, readable brand/gender, FRAME SHAPE column that mixes rim type and shape,
# size embedded as first 2 chars of "size+colour". Self-contained handler
# returning the same (df, unmapped, skipped) shape as load_single_catalog.

_DERIGO_BRAND_MAP = {
    "police": "Police",
    "furla": "Furla",
    "just cavalli": "Just Cavalli",
}

# FRAME SHAPE column carries BOTH rim type and (occasionally) actual shape.
_DERIGO_RIM_MAP = {
    "FULL-FRAME": "Full rim",
    "FULL FRAME": "Full rim",
    "RIMLESS": "Rimless",
    "HALF-FRAME": "Half rim",
    "HALF FRAME": "Half rim",
    "SEMI-RIMLESS": "Half rim",
}
_DERIGO_SHAPE_MAP = {
    "SQUARE": "Square",
    "RECTANGULAR": "Rectangular",
    "ROUND": "Round",
    "GEOMETRIC": "Extravagant",
    "CAT": "Cat Eye",
    "CAT-EYE": "Cat Eye",
    "NAVIGATOR": "Pilot",
    "PILOT": "Pilot",
    "SHIELD": "Single lens",
    "OVAL": "Oval / Elipse",
    "BUTTERFLY": "Butterfly",
    "BROWLINE": "Browline",
    "PANTOS": "Panthos / Tea cup",
    "PANTHOS": "Panthos / Tea cup",
}

_DERIGO_GENDER_MAP = {
    "WOMAN": "Woman",
    "MAN": "Man",
    "UNISEX": "Man|Woman",
    "JUNIOR": "Child",
    "KIDS": "Child",
    "CHILD": "Child",
}

_DERIGO_ORIGIN_MAP = {
    "CN": "China", "BD": "Bangladesh", "VN": "Vietnam",
    "KH": "Cambodia", "IT": "Italy", "JP": "Japan", "FR": "France",
}


def _derigo_material(raw):
    """Map FRONT MATERIAL free-text to Plastic / Metal / Titanium."""
    s = str(raw or "").strip()
    low = s.lower()
    if not s or low == "nan":
        return "", False
    if "titanium" in low:
        return "Titanium", True
    if any(k in low for k in ("acetate", "acetato", "injected", "pet", "rpet", "nylon", "tritan", "renew", "plastic")):
        return "Plastic", True
    if any(k in low for k in ("steel", "metal", "monel", "alumin", "alloy", "bronze")):
        return "Metal", True
    return s, False  # unknown — keep original, flag


def _derigo_lens_base(raw):
    """Parse LENS BASE -> (filter_category, is_polarized, is_photochromic).
    Examples: 'CATEGORIA *3', 'CATEGORIA *3 POLARIZZANTE',
    'CATEGORIA *1-3 FOTOCROMATICA'. 'BASE 2'/'BASE 4' are base curves -> no cat."""
    s = str(raw or "").strip().upper()
    if not s or s == "NAN":
        return "", False, False
    polarized = "POLARIZZANTE" in s
    photochromic = "FOTOCROMATIC" in s
    cat = ""
    m = re.search(r"CATEGORIA\s*\*?\s*(\d)\s*-\s*(\d)", s)
    if m:
        lo, hi = sorted([m.group(1), m.group(2)], key=int)
        cat = f"Category range {lo} - {hi}"
    else:
        m2 = re.search(r"CATEGORIA\s*\*?\s*(\d)", s)
        if m2:
            cat = f"Category {m2.group(1)}"
    return cat, polarized, photochromic


def _derigo_lens_material(raw):
    s = str(raw or "").strip()
    low = s.lower()
    if not s or low == "nan":
        return "", False
    if "cr39" in low or "cr 39" in low:
        return "CR 39", True
    if "nylon" in low:
        return "Nylon", True
    if "policarbon" in low or "polycarbon" in low:
        return "Polycarbonate", True
    if "triacetato" in low or "tca" in low or "tritan" in low or "copoliestere" in low or "poliestere" in low:
        return "Plastic", True
    return s, False  # unknown — keep original, flag


def _load_derigo(df):
    unmapped = set()
    skipped = set()
    rows = []

    for _, src in df.iterrows():
        barcode = str(src.get("EAN", "")).strip()
        if not barcode or barcode.lower() == "nan":
            continue
        join_key = re.sub(r"\.0$", "", barcode).lstrip("0")
        if not join_key or join_key == "nan":
            continue

        out = {"Barcode": barcode, "join_key": join_key}

        # ---- Brand ----
        brand_raw = str(src.get("Brand", "")).strip()
        brand_norm = re.sub(r"\s+", " ", brand_raw).strip().lower()
        brand = _DERIGO_BRAND_MAP.get(brand_norm, brand_raw)
        if brand_norm and brand_norm not in _DERIGO_BRAND_MAP:
            unmapped.add(f"Derigo -> Brand: '{brand_raw}'")
        out["Brand"] = brand
        out["Manufacturer"] = brand

        # ---- Type ----
        so = str(src.get("Sun/optical", "")).strip().lower()
        if so.startswith("sun"):
            out["Glasses_type"] = "Sunglasses"
        elif so.startswith("opt"):
            out["Glasses_type"] = "Frames"
        else:
            out["Glasses_type"] = ""

        # ---- size+colour -> lens width + colour code ----
        sc = str(src.get("size+colour", "")).strip()
        m = re.match(r"(\d{2})(.*)", sc)
        if m:
            size = m.group(1)
            color_code = m.group(2).strip()
        else:
            size = ""
            color_code = sc
        out["Glasses_size_lens_width"] = size
        out["Combination"] = size
        out["Glasses_color_code"] = color_code

        # ---- Other dimensions ----
        out["Glasses_size_bridge"] = _marcolin_round(src.get("BRIDGE LENGHT"))
        out["Glasses_size_temple_length"] = _marcolin_round(src.get("TEMPLE LENGHT"))
        out["Glasses_size_lens_height"] = _marcolin_round(src.get("LENS HIGHT"))

        # ---- FRAME SHAPE: split into rim type vs shape ----
        fs = str(src.get("FRAME SHAPE", "")).strip()
        fs_u = fs.upper()
        out["Glasses_frame_type"] = ""
        out["Glasses_shape"] = ""
        if fs and fs_u != "NAN":
            if fs_u in _DERIGO_RIM_MAP:
                out["Glasses_frame_type"] = _DERIGO_RIM_MAP[fs_u]
            elif fs_u in _DERIGO_SHAPE_MAP:
                out["Glasses_shape"] = _DERIGO_SHAPE_MAP[fs_u]
            else:
                out["Glasses_shape"] = fs  # keep original, flag
                unmapped.add(f"Derigo -> FRAME SHAPE: '{fs}'")

        # ---- Materials ----
        mat, ok = _derigo_material(src.get("FRONT MATERIAL"))
        out["Glasses_main_material"] = mat
        if mat and not ok:
            unmapped.add(f"Derigo -> Glasses_main_material: '{src.get('FRONT MATERIAL')}'")
        lm, ok2 = _derigo_lens_material(src.get("LENS MATERIAL"))
        out["Glasses_lens_material"] = lm
        if lm and not ok2:
            unmapped.add(f"Derigo -> Glasses_lens_material: '{src.get('LENS MATERIAL')}'")

        # ---- Colours ----
        frame_col = str(src.get("COLOR DESCRIPTION", "")).strip()
        lens_col = str(src.get("LENS COLOR", "")).strip()
        if frame_col and frame_col.lower() != "nan":
            res = classify_color(frame_col, "frame")
            out["Frame_Colour"] = res if res else frame_col
            out["Temple_Colour"] = res if res else frame_col
            if not res:
                unmapped.add(f"Derigo -> Frame_Colour: '{frame_col}'")
        else:
            out["Frame_Colour"] = ""
            out["Temple_Colour"] = ""
        if lens_col and lens_col.lower() != "nan":
            res = classify_color(lens_col, "lens")
            out["Glasses_lens_Colour"] = res if res else lens_col
            if not res:
                unmapped.add(f"Derigo -> Glasses_lens_Colour: '{lens_col}'")
        else:
            out["Glasses_lens_Colour"] = ""

        # ---- Filter category + lens effect (both from LENS BASE) ----
        cat, pol_lb, photo_lb = _derigo_lens_base(src.get("LENS BASE"))
        out["Sunglasses_filter"] = cat

        eff = set()
        if str(src.get("POLARIZED (YES/NO)", "")).strip().upper() == "YES" or pol_lb:
            eff.add("Polarized")
        if photo_lb:
            eff.add("Photochromic")
        out["Glasses_lens_effect"] = "|".join(sorted(eff))

        # ---- Gender ----
        g = str(src.get("GENDER", "")).strip()
        if not g or g.upper() == "NAN":
            out["Glasses_gendre"] = ""
        elif g.upper() in _DERIGO_GENDER_MAP:
            out["Glasses_gendre"] = _DERIGO_GENDER_MAP[g.upper()]
        else:
            out["Glasses_gendre"] = g  # keep original, flag
            unmapped.add(f"Derigo -> Glasses_gendre: '{g}'")

        # ---- RX: no source column ----
        out["SunGlasses_RX_lenses"] = ""

        # ---- Origin ----
        origin = str(src.get("COUNTY OF ORIGIN", "")).strip().upper()
        out["Item_origin_country"] = _DERIGO_ORIGIN_MAP.get(
            origin, origin if origin and origin != "NAN" else ""
        )

        # ---- Model + name ----
        model = str(src.get("model", "")).strip()
        out["Extracted_Model"] = model
        out["Extracted_Color"] = color_code
        name_parts = [p for p in (brand, model, color_code) if p and p.lower() != "nan"]
        out["Assembled_Name"] = " ".join(name_parts)

        # ---- Clip-on: De Rigo has no explicit clip column, but an OPTICAL
        # frame that carries a sun filter category in LENS BASE is an optical
        # frame bundled with a (magnetic) sun clip-on. Polarized -> ' p' variant.
        clip = ""
        if out["Glasses_type"] == "Frames" and cat:
            clip = "Magnetic sun clip-on p" if pol_lb else "Magnetic sun clip-on"
        out["Extracted_Clip_on"] = clip
        out["Clip_on_Alert"] = False

        out["Producing_company"] = "Derigo"
        rows.append(out)

    return pd.DataFrame(rows), unmapped, skipped


# ==========================================================================
# ALPINA — German sports-eyewear master data
# ==========================================================================
# All sunglasses (Sportbrillen) and ski goggles (Skibrillen); no optical
# frames. German values throughout. Very complete: dimensions, filter
# category, materials, lens effects (Technologien column), gender, origin.

_ALPINA_SHAPE_MAP = {
    "RECHTECKIG": "Rectangular",
    "RUND": "Round",
    "OVAL": "Oval / Elipse",
    "PILOT": "Pilot",
    "MONOSCHEIBE": "Single lens",
    "SPORTLICH (WRAP)": "Extravagant",
    "SPORTLICH": "Extravagant",
    "WRAP": "Extravagant",
    "SCHMETTERLING": "Butterfly",
    "QUADRATISCH": "Square",
    "KATZENAUGE": "Cat Eye",
}

_ALPINA_RIM_MAP = {
    "VOLLRAND": "Full rim",
    "HALBRAND": "Half rim",
    "RANDLOS": "Rimless",
}


def _alpina_material(raw):
    """German frame/lens material -> Plastic / Metal / Titanium / etc."""
    s = str(raw or "").strip()
    low = s.lower()
    if not s or low in ("nan", "(blank)"):
        return None
    if "titan" in low:
        return "Titanium"
    if "neusilber" in low or "metall" in low or "alumin" in low or "stahl" in low:
        return "Metal"
    if "polycarbonat" in low or "polykarbonat" in low:
        return "Polycarbonate"
    if "nylon" in low:
        return "Nylon"
    if "triacetat" in low or "tac" in low:
        return "Plastic"
    if any(k in low for k in ("polyamid", "tr 90", "tr90", "tpee", "tpr", "tpu",
                              "polyester", "elastomer", "polyurethan", "nxt", "kunststoff", "pa")):
        return "Plastic"
    return None


def _alpina_filter(raw):
    """'cat. 3' -> 'Category 3'; 'cat. 1-cat. 3' -> 'Category range 1 - 3'."""
    s = str(raw or "").strip().lower()
    if not s or s in ("nan", "(blank)"):
        return None
    nums = re.findall(r"(\d)", s)
    if not nums:
        return None
    if len(nums) >= 2 and ("-" in s or "bis" in s):
        lo, hi = sorted([nums[0], nums[-1]], key=int)
        return f"Category range {lo} - {hi}"
    return f"Category {nums[0]}"


def _load_alpina(df):
    unmapped = set()
    skipped = set()
    rows = []

    df = df.copy()
    df.columns = df.columns.astype(str).str.replace("\xa0", " ", regex=False).str.strip()

    def col(*names):
        for n in names:
            for c in df.columns:
                if n.lower() == c.lower():
                    return c
        for n in names:  # fuzzy contains
            for c in df.columns:
                if n.lower() in c.lower():
                    return c
        return None

    C = {
        "ean": col("Europäische Artikelnummer EAN", "EAN"),
        "type": col("Bezeichnung PL2"),
        "prod": col("Materialnummer Produkt"),
        "kurz": col("Materialkurztext"),
        "farbe": col("Farben Merkmalswerteliste"),
        "gender": col("Geschlecht"),
        "form": col("Form"),
        "rim": col("Rahmentyp"),
        "filter": col("Filterkategorie"),
        "frame_col": col("Rahmen Farbe"),
        "temple_col": col("Bügel Farbe"),
        "lens_col": col("Scheibenfarben Merkmalswerteli", "Scheibenfarben"),
        "frame_mat": col("Rahmen Material"),
        "lens_mat": col("Scheibe Material"),
        "origin": col("Ursprungsland des Materials"),
        "weight": col("Nettogewicht"),
        "lens_w": col("Glasbreite"),
        "bridge": col("Nasenstegbreite"),
        "temple": col("Bügellänge"),
        "lens_h": col("Glashöhe"),
        "mirror": col("Scheibe Verspiegelung (außen)", "Verspiegelung"),
        "tech": col("Technologien"),
    }

    def g(row, key):
        c = C.get(key)
        return str(row.get(c, "")).strip() if c else ""

    for _, row in df.iterrows():
        barcode = g(row, "ean")
        if not barcode or barcode.lower() == "nan":
            continue
        join_key = re.sub(r"\.0$", "", barcode).lstrip("0")
        if not join_key or join_key == "nan":
            continue

        out = {"Barcode": barcode, "join_key": join_key}
        out["Brand"] = "Alpina"
        out["Manufacturer"] = "Alpina"

        # Type: Skibrillen -> Sport glasses, everything else -> Sunglasses
        t = g(row, "type").lower()
        out["Glasses_type"] = "Sport glasses" if "ski" in t else "Sunglasses"
        is_sun = True  # both sunglasses and ski goggles carry lens data

        # Dimensions (strip " mm", round)
        out["Glasses_size_lens_width"] = _marcolin_round(g(row, "lens_w")) or None
        out["Combination"] = _marcolin_round(g(row, "lens_w")) or None
        out["Glasses_size_bridge"] = _marcolin_round(g(row, "bridge")) or None
        out["Glasses_size_temple_length"] = _marcolin_round(g(row, "temple")) or None
        out["Glasses_size_lens_height"] = _marcolin_round(g(row, "lens_h")) or None

        # Shape
        form = g(row, "form")
        if not form or form.lower() == "nan":
            out["Glasses_shape"] = None
        elif form.upper() in _ALPINA_SHAPE_MAP:
            out["Glasses_shape"] = _ALPINA_SHAPE_MAP[form.upper()]
        else:
            out["Glasses_shape"] = form
            unmapped.add(f"Alpina -> Glasses_shape: '{form}'")

        # Rim
        rim = g(row, "rim")
        if not rim or rim.lower() == "nan":
            out["Glasses_frame_type"] = None
        elif rim.upper() in _ALPINA_RIM_MAP:
            out["Glasses_frame_type"] = _ALPINA_RIM_MAP[rim.upper()]
        else:
            out["Glasses_frame_type"] = rim
            unmapped.add(f"Alpina -> Glasses_frame_type: '{rim}'")

        # Materials
        out["Glasses_main_material"] = _alpina_material(g(row, "frame_mat"))
        if g(row, "frame_mat") and out["Glasses_main_material"] is None and g(row, "frame_mat").lower() != "nan":
            unmapped.add(f"Alpina -> Glasses_main_material: '{g(row, 'frame_mat')}'")
        out["Glasses_lens_material"] = _alpina_material(g(row, "lens_mat"))

        # Colours
        for src_key, out_col, ctype in [
            ("frame_col", "Frame_Colour", "frame"),
            ("temple_col", "Temple_Colour", "frame"),
            ("lens_col", "Glasses_lens_Colour", "lens"),
        ]:
            v = g(row, src_key)
            if v and v.lower() != "nan":
                res = classify_color(v, ctype)
                out[out_col] = res if res else v
                if not res:
                    unmapped.add(f"Alpina -> {out_col}: '{v}'")
            else:
                out[out_col] = None

        # Filter category
        out["Sunglasses_filter"] = _alpina_filter(g(row, "filter"))

        # Lens effect (from Technologien + Verspiegelung + lens colour/material)
        tech = g(row, "tech").lower()
        lens_col_v = g(row, "lens_col").lower()
        lens_mat_v = g(row, "lens_mat").lower()
        eff = set()
        if "polaris" in tech or "polaris" in lens_col_v or "pola" in lens_mat_v:
            eff.add("Polarized")
        if "mirror" in tech or "mirror" in lens_col_v or g(row, "mirror").lower() == "ja":
            eff.add("Mirror")
        if "photochrom" in tech or "selbsttön" in tech or "varioflex" in lens_col_v or "vario" in tech:
            eff.add("Photochromic")
        if "verlauf" in lens_col_v or "gradient" in lens_col_v or "gradient" in tech:
            eff.add("Gradient")
        out["Glasses_lens_effect"] = "|".join(sorted(eff)) if eff else None

        # Gender
        gd = g(row, "gender").lower()
        genders = set()
        if "herren" in gd or "jungen" in gd:
            genders.add("Man")
        if "damen" in gd or "mädchen" in gd or "madchen" in gd:
            genders.add("Woman")
        if "unisex" in gd:
            genders.update({"Man", "Woman"})
        # Kids: Jungen/Mädchen -> Child
        if "jungen" in gd or "mädchen" in gd or "madchen" in gd or "kinder" in gd:
            out["Glasses_gendre"] = "Child"
        elif genders:
            out["Glasses_gendre"] = "|".join(sorted(genders))
        else:
            out["Glasses_gendre"] = None

        # Origin (e.g. 'CN: VR China' -> China)
        origin_raw = g(row, "origin")
        code = origin_raw.split(":")[0].strip().upper() if origin_raw else ""
        origin_map = dict(_MARCOLIN_ORIGIN_MAP)
        origin_map.update({"TW": "Taiwan", "DE": "Germany"})
        out["Item_origin_country"] = origin_map.get(code, None)

        # Weight (grams)
        w = g(row, "weight")
        if w and w.lower() != "nan":
            try:
                out["Glasses_weight_g"] = str(int(round(float(w.replace(",", ".")))))
            except Exception:
                pass

        # Model / colour code / name
        model = g(row, "prod")
        colour_code = g(row, "farbe")
        out["Extracted_Model"] = model
        out["Extracted_Color"] = colour_code
        out["Glasses_color_code"] = colour_code
        kurz = g(row, "kurz")
        out["Assembled_Name"] = f"Alpina {kurz}".strip() if kurz and kurz.lower() != "nan" else f"Alpina {model}".strip()

        out["Extracted_Clip_on"] = ""
        out["Clip_on_Alert"] = False
        out["Producing_company"] = "Alpina"
        rows.append(out)

    return pd.DataFrame(rows), unmapped, skipped


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

    # Tom Ford catalogue (Marcolin family, Tom Ford column names). Upload as
    # "marcolin"; detected by its signature columns.
    if mfg_name == "marcolin" and {"MAIN MATERIAL", "SIZE/COLOR", "MADE IN"}.issubset(set(df.columns)):
        return _load_marcolin_tomford(df)

    # Safilo catalog export (2nd Safilo format: optical frames / opt+clip-on /
    # sunglasses) — detected by its signature columns. Additive top-up handler.
    if mfg_name == "safilo" and {"Division Description", "Color Code Description"}.issubset(set(df.columns)):
        return _load_safilo_catalog(df)

    # De Rigo (Police / Furla / Just Cavalli) — dedicated format handler.
    if mfg_name == "derigo":
        return _load_derigo(df)

    # Thélios (LVMH: Celine, Dior, Bulgari, Loewe, Fendi, ...) — dedicated handler.
    if mfg_name == "thelios":
        return _load_thelios(df)

    # Alpina (German sports eyewear) — dedicated handler.
    if mfg_name == "alpina":
        return _load_alpina(df)

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
                        final_values.add(first_shape)  # keep original, flag for review
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
                    # No keyword match — keep the original value so the cell
                    # isn't empty and the validator can flag it for manual fix.
                    unmapped_values.add(f"{mfg.title()} -> {col_name}: '{raw_val}'")
                    final_values.add(raw_val)

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
                        # Unknown value — keep the original so the cell isn't
                        # empty and the validator can flag it for manual fix.
                        unmapped_values.add(f"{mfg.title()} -> {col_name}: '{part}'")
                        final_values.add(part)
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


def record_ingest(engine, mfg_key, row_count=None):
    """Record that a manufacturer's catalogue was just processed, into the
    `ingest_log` table (one row per manufacturer, upserted). Used by both the
    admin app and the headless auto-ingest so the 'last updated' panel reflects
    every import. Never raises — a logging failure must not fail an ingest."""
    from datetime import datetime, timezone
    try:
        ts = datetime.now(timezone.utc).isoformat(timespec="seconds")
        try:
            existing = pd.read_sql_table("ingest_log", con=engine)
        except Exception:
            existing = pd.DataFrame(columns=["manufacturer", "last_updated", "rows"])
        existing = existing[
            existing["manufacturer"].astype(str).str.lower() != str(mfg_key).lower()
        ]
        new_row = pd.DataFrame([{
            "manufacturer": str(mfg_key),
            "last_updated": ts,
            "rows": int(row_count) if row_count is not None else None,
        }])
        combined = pd.concat([existing, new_row], ignore_index=True)
        combined.to_sql("ingest_log", con=engine, if_exists="replace", index=False)
    except Exception:
        pass


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
