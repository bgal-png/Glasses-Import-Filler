# -*- coding: utf-8 -*-
"""
One-shot fill script for Tom Ford / Max Mara / Moncler data delivered in the
"MATERIAL INFO AMBG" XLSX format (different structure than regular Marcolin
files — uses Italian conventions like REXABLE=SI, separate FRONT/TEMPLE
material columns, COLOR DESCRIPTION as "front / lens" pairs).

Reads the source file in memory, translates each row to the canonical global
column names + values, then fills a target template by matching on barcode.
Does NOT touch master_catalog or any Supabase tables — purely local.

Usage:
    python scripts/oneshot_fill_tomford.py <source.xlsx> <target.xlsx> <output.xlsx>
"""
from __future__ import annotations

import os
import re
import sys

import pandas as pd

# Make repo root importable
_REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, _REPO)

from dictionaries import (  # noqa: E402
    TARGET_MAPPING,
    BRAND_USABLE_MAP,
    BRAND_GLASSES_CONTAIN,
    FACE_SHAPE_MAP,
    classify_color,
)

# ----------------------------------------------------------------------
# Translation tables tailored to the MATERIAL INFO AMBG format
# ----------------------------------------------------------------------
BRAND_CODE_MAP = {
    "FT": "Tom Ford",
    "MM": "Max Mara",
    "MO": "Max&Co.",
}

SHAPE_MAP = {
    "ROUND": "Round",
    "RECTANGULAR": "Rectangular",
    "SQUARE": "Square",
    "OVAL": "Oval / Elipse",
    "PILOT": "Pilot",
    "NAVIGATOR": "Pilot",
    "BUTTERFLY": "Butterfly",
    "CAT": "Cat Eye",
    "GEOMETRIC": "Extravagant",
    "SHIELD": "Single lens",
}

GENDER_MAP = {"F": "Woman", "M": "Man", "U": "Man|Woman"}

ORIGIN_MAP = {"IT": "Italy", "CN": "China", "JP": "Japan", "FR": "France"}

RIM_MAP = {
    "FULL RIM": "Full rim",
    "SEMIRIMLESS": "Half rim",
    "HALF RIM": "Half rim",
    "RIMLESS": "Rimless",
}

LENSES_CAT_MAP = {
    "0": "Category 0",
    "1": "Category 1",
    "2": "Category 2",
    "3": "Category 3",
    "4": "Category 4",
    "1/3": "Category range 1 - 3",
    "0/3": "Category range 0 - 3",
    "2/3": "Category range 2 - 3",
    "3P": "Category 3",  # polarized recorded separately in lens effect
}

MATERIAL_MAP = {
    "ACETATE": "Plastic",
    "INJECTED": "Plastic",
    "VARNISHED METAL": "Metal",
    "MAGNESIUM": "Metal",
    "METAL": "Metal",
    "PLASTIC": "Plastic",
    "NYLON": "Plastic",
    "TITANIUM": "Titanium",
}


def _round_dim(v: object) -> str:
    s = str(v if v is not None else "").strip()
    if not s or s.lower() == "nan":
        return ""
    try:
        clean = re.sub(r"[^\d,.-]", "", s).replace(",", ".")
        return str(int(round(float(clean)))) if clean else s
    except Exception:
        return s


def translate_source_row(src_row: pd.Series) -> tuple[str, dict] | None:
    """One source row -> (barcode, dict-of-global-values). Returns None if no barcode."""
    barcode = str(src_row.get("EAN/UPC CODE", "")).strip()
    if not barcode or barcode.lower() == "nan":
        return None

    out: dict = {}

    # ---- Brand & Manufacturer ----
    brand_code = str(src_row.get("BRAND", "")).strip().upper()
    brand = BRAND_CODE_MAP.get(brand_code, brand_code)
    out["Brand"] = brand
    out["Manufacturer"] = brand

    # ---- Glasses type (from DESCRIPTION text) ----
    desc = str(src_row.get("DESCRIPTION", "")).strip().lower()
    if "sunglass" in desc:
        out["Glasses_type"] = "Sunglasses"
    elif "frame" in desc:
        out["Glasses_type"] = "Frames"
    else:
        out["Glasses_type"] = ""

    # ---- Dimensions ----
    lens_w = _round_dim(src_row.get("SIZE"))
    out["Glasses_size_lens_width"] = lens_w
    out["Glasses_size_bridge"] = _round_dim(src_row.get("NOSE-BRIDGE SIZE"))
    out["Glasses_size_temple_length"] = _round_dim(src_row.get("TEMPLE LENGHT"))
    out["Glasses_size_lens_height"] = _round_dim(src_row.get("B Measurement"))
    out["Combination"] = lens_w  # Combination is just lens width

    # ---- Shape, rim, flex, gender ----
    shape = str(src_row.get("FORM DESCRIPTION", "")).strip().upper()
    out["Glasses_shape"] = SHAPE_MAP.get(shape, "")

    rim = str(src_row.get("RIM DESCRIPTION", "")).strip().upper()
    out["Glasses_frame_type"] = RIM_MAP.get(rim, "")

    flex = str(src_row.get("FLEX", "")).strip().upper()
    out["Glasses_other_info"] = "Flex" if flex == "SI" else ""

    g = str(src_row.get("GENDER", "")).strip().upper()
    out["Glasses_gendre"] = GENDER_MAP.get(g, "")

    # ---- Materials ----
    front_mat = str(src_row.get("DESCRIPTION FRONT", "")).strip().upper()
    out["Glasses_main_material"] = MATERIAL_MAP.get(front_mat, "")

    # ---- RX & origin ----
    rx = str(src_row.get("REXABLE", "")).strip().upper()
    out["SunGlasses_RX_lenses"] = "Yes" if rx == "SI" else ""

    origin = str(src_row.get("ORIGIN", "")).strip().upper()
    out["Item_origin_country"] = ORIGIN_MAP.get(origin, origin)

    # ---- Lens filter category (drop trailing P for polarized) ----
    cat = str(src_row.get("LENSES CATEGORY", "")).strip().upper()
    out["Sunglasses_filter"] = LENSES_CAT_MAP.get(cat, "")

    # ---- Lens effect (polarized / photochromic) ----
    eff: set[str] = set()
    lens_desc = str(src_row.get("LENSES DESCRIPTION", "")).strip().upper()
    if "POLAR" in lens_desc or cat.endswith("P"):
        eff.add("Polarized")
    if "PHOTO" in lens_desc:
        eff.add("Photochromic")
    out["Glasses_lens_effect"] = "|".join(sorted(eff))

    # ---- Colors (format: "front color / lens color [modifier]") ----
    color_desc = str(src_row.get("COLOR DESCRIPTION", "")).strip()
    if color_desc and color_desc.lower() != "nan":
        parts = [p.strip() for p in color_desc.split("/")]
        front_color = parts[0] if parts else ""
        lens_color_part = parts[1] if len(parts) > 1 else ""
        if front_color:
            res = classify_color(front_color, "frame")
            out["Frame_Colour"] = res
            out["Temple_Colour"] = res
        if lens_color_part:
            out["Glasses_lens_Colour"] = classify_color(lens_color_part, "lens")

    # ---- Lens material (inferred from LENSES DESCRIPTION since source has no
    # explicit lens-material column). Common Tom Ford / Max Mara / Moncler
    # defaults. Frames have no lens material — leave empty for those.
    if "Sunglasses" in out["Glasses_type"]:
        if "POLAR" in lens_desc:
            out["Glasses_lens_material"] = "Polar CR 39"
        elif "PHOTO" in lens_desc:
            out["Glasses_lens_material"] = "Plastic"
        elif lens_desc == "DEMO":
            out["Glasses_lens_material"] = ""
        else:  # NORMAL or blank
            out["Glasses_lens_material"] = "CR 39"
    else:
        out["Glasses_lens_material"] = ""

    # ---- Clip-on ----
    clipon = str(src_row.get("CLIPON", "")).strip()
    out["Extracted_Clip_on"] = "Sun clip-on" if "Included" in clipon else ""

    # ---- Weights ----
    nw = str(src_row.get("NET WEIGHT", "")).strip()
    if nw and nw.lower() != "nan":
        try:
            grams = round(float(nw) * 1000)  # source is kg
            out["Glasses_weight_g"] = str(grams)
        except Exception:
            pass

    # ---- Stamping ----
    out["Producing_company"] = "Marcolin"  # all 3 brands are Marcolin-distributed

    # ---- Model, color code, assembled name ----
    # In this format the SKU has the size baked in as a prefix
    # (e.g. SIZE=52, SKU=5201D -> real color code is "01D"). Strip the
    # size prefix when present to get the actual color code.
    model = str(src_row.get("MODEL", "")).strip()
    sku = str(src_row.get("SKU", "")).strip()
    size_str = str(src_row.get("SIZE", "")).strip()
    color_code = sku
    if size_str and sku.startswith(size_str):
        color_code = sku[len(size_str):]
    out["Extracted_Model"] = model
    out["Extracted_Color"] = color_code
    out["Glasses_color_code"] = color_code
    name_parts = [p for p in (brand, model, color_code) if p and p.lower() != "nan"]
    out["Assembled_Name"] = " ".join(name_parts)

    out["Barcode"] = barcode
    return barcode, out


def build_master_db(src_file: str) -> pd.DataFrame:
    """Read source file, return DataFrame indexed by clean barcode (lstripped of zeros)."""
    df_src = pd.read_excel(src_file, dtype=str, engine="openpyxl")
    df_src.columns = df_src.columns.astype(str).str.strip()

    rows = []
    for _, src_row in df_src.iterrows():
        result = translate_source_row(src_row)
        if result is None:
            continue
        barcode, data = result
        clean_bc = re.sub(r"\.0$", "", barcode.strip()).lstrip("0")
        if not clean_bc:
            continue
        data["join_key"] = clean_bc
        rows.append(data)

    df = pd.DataFrame(rows)
    df.set_index("join_key", inplace=True)
    return df


def fill_target(target_file: str, master_db: pd.DataFrame, output_file: str) -> None:
    target_df = pd.read_excel(target_file, dtype=str, engine="openpyxl")
    target_df.columns = (
        target_df.columns.astype(str)
        .str.replace("\n", " ", regex=False)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )

    # Detect "Glasses contain" column variant (with/without space before 84)
    if "Glasses contain ID:84" in target_df.columns:
        contain_col = "Glasses contain ID:84"
    elif "Glasses contain ID: 84" in target_df.columns:
        contain_col = "Glasses contain ID: 84"
    else:
        contain_col = "Glasses contain ID:84"
        target_df[contain_col] = ""

    target_bc_col = TARGET_MAPPING.get("Barcode", "Barcode")
    if target_bc_col not in target_df.columns:
        raise SystemExit(f"Target file missing barcode column '{target_bc_col}'")

    # Ensure all expected target columns exist
    for global_col, target_col in TARGET_MAPPING.items():
        if isinstance(target_col, list):
            for tc in target_col:
                if tc not in target_df.columns:
                    target_df[tc] = ""
        else:
            if target_col not in target_df.columns:
                target_df[target_col] = ""

    matched = 0
    unmatched: list[str] = []

    for idx, row in target_df.iterrows():
        raw_bc = str(row[target_bc_col]).strip()
        clean_bc = re.sub(r"\.0$", "", raw_bc).lstrip("0")
        if clean_bc not in master_db.index:
            if raw_bc and raw_bc.lower() != "nan":
                unmatched.append(raw_bc)
            continue
        matched += 1

        master_row = master_db.loc[clean_bc]
        if isinstance(master_row, pd.DataFrame):
            master_row = master_row.iloc[0]

        g_type = str(master_row.get("Glasses_type", "")).strip()
        is_frames = g_type == "Frames"
        lens_skip = {
            "Glasses_lens_Colour", "Glasses_lens_material",
            "Sunglasses_filter", "Glasses_lens_effect",
            "SunGlasses_RX_lenses",
        }

        # ---- Standard TARGET_MAPPING fills ----
        for global_col, target_col in TARGET_MAPPING.items():
            if global_col == "Barcode":
                continue
            if is_frames and global_col in lens_skip:
                continue
            if global_col not in master_db.columns:
                continue
            val = master_row.get(global_col, "")
            if pd.notna(val) and str(val).strip():
                val_str = str(val).strip()
                if isinstance(target_col, list):
                    for tc in target_col:
                        target_df.at[idx, tc] = val_str
                else:
                    target_df.at[idx, target_col] = val_str

        # ---- Constants ----
        target_df.at[idx, "Items type ID: 20"] = "Glasses"
        target_df.at[idx, "Items packing ID: 21"] = "Basic"

        assembled = str(master_row.get("Assembled_Name", "")).strip()
        raw_mat = str(master_row.get("Glasses_main_material", "")).strip().lower()

        # ---- Meta description / Item description / HS code ----
        if "Sunglasses" in g_type:
            target_df.at[idx, "Meta description"] = f"Sunglasses {assembled}"
            target_df.at[idx, "HS Code"] = "90041091"
            has_p = "plastic" in raw_mat
            has_m = "metal" in raw_mat
            if has_p and has_m:
                target_df.at[idx, "Item description"] = "Sunglasses, mixed plastic and metal frame"
            elif has_p:
                target_df.at[idx, "Item description"] = "Sunglasses, plastic frame"
            elif has_m:
                target_df.at[idx, "Item description"] = "Sunglasses, metal frame"
            target_df.at[idx, "UV filter ID: 60"] = "400"
        elif "Frames" in g_type:
            target_df.at[idx, "Meta description"] = f"Eyeglasses {assembled}"
            target_df.at[idx, "Item description"] = "Eyeglasses"
            if "plastic" in raw_mat:
                target_df.at[idx, "HS Code"] = "90031100"
            elif "metal" in raw_mat:
                target_df.at[idx, "HS Code"] = "90031900"

        # ---- Glasses usable ----
        usable: set[str] = set()
        brand_lower = str(master_row.get("Brand", "")).strip().lower()
        if brand_lower in BRAND_USABLE_MAP:
            usable.add(BRAND_USABLE_MAP[brand_lower])
        eff_val = str(master_row.get("Glasses_lens_effect", "")).strip()
        if "Sunglasses" in g_type:
            if "Polarized" in eff_val:
                usable.add("Driving glasses")
            else:
                usable.add("Common use")
        if usable:
            target_df.at[idx, "Glasses usable ID: 51"] = "|".join(sorted(usable))

        # ---- Face shape recommendation ----
        shape_raw = str(master_row.get("Glasses_shape", "")).strip()
        if shape_raw and shape_raw.lower() != "nan":
            faces: set[str] = set()
            for s in shape_raw.split("|"):
                for shape_key, face_val in FACE_SHAPE_MAP.items():
                    if shape_key.lower() == s.strip().lower():
                        for face in face_val.split("|"):
                            faces.add(face)
            if faces:
                target_df.at[idx, "Glasses for your face shape ID:94"] = "|".join(sorted(faces))

        # ---- Glasses contain (per brand x type, from BRAND_GLASSES_CONTAIN dict) ----
        type_key = "Sunglasses" if "Sunglasses" in g_type else "Frames"
        contain_str = ""
        entry = BRAND_GLASSES_CONTAIN.get(brand_lower)
        if entry:
            contain_str = entry.get(type_key, "")
        clip_val = str(master_row.get("Extracted_Clip_on", "")).strip()

        final_contain: list[str] = []
        if contain_str:
            final_contain.extend(contain_str.split("|"))
        if clip_val and clip_val.lower() not in ("nan", ""):
            final_contain.append(clip_val)
        if final_contain:
            uniq = {item.strip().lower(): item.strip() for item in final_contain if item.strip()}
            ordered = []
            if "original glasses case" in uniq:
                ordered.append("Original glasses case")
                del uniq["original glasses case"]
            if "cleaning cloth" in uniq:
                ordered.append("Cleaning cloth")
                del uniq["cleaning cloth"]
            ordered.extend(sorted(uniq.values()))
            target_df.at[idx, contain_col] = "|".join(ordered)

    target_df.to_excel(output_file, index=False, engine="openpyxl")

    print()
    print(f"Matched: {matched} / {len(target_df)} rows from source")
    if unmatched:
        print(f"Unmatched: {len(unmatched)} barcodes (these aren't in the source file)")
        for bc in unmatched[:20]:
            print(f"  - {bc}")
        if len(unmatched) > 20:
            print(f"  … and {len(unmatched) - 20} more")
    print()
    print(f"Output written to: {output_file}")


def main():
    if len(sys.argv) != 4:
        print("Usage: python scripts/oneshot_fill_tomford.py <source.xlsx> <target.xlsx> <output.xlsx>")
        sys.exit(1)

    src, tgt, out = sys.argv[1], sys.argv[2], sys.argv[3]
    print(f"Source: {src}")
    print(f"Target: {tgt}")
    print(f"Output: {out}")
    print()
    print("Building in-memory master_db from source…")
    master_db = build_master_db(src)
    print(f"  Translated {len(master_db)} rows.")
    print(f"  Brands: {sorted(master_db['Brand'].unique().tolist())}")
    print()
    print("Filling target template…")
    fill_target(tgt, master_db, out)


if __name__ == "__main__":
    main()
