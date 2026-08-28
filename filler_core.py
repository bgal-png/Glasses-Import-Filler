# -*- coding: utf-8 -*-
"""
Pure (UI-free) auto-filler engine.

No Streamlit, no Qt — so it can be driven by:
  - app_manufacturer.py   (Streamlit web app)
  - desktop/              (PySide6 desktop app)
  - headless scripts / tests

Public API
----------
    read_target_file(path_or_buffer, is_csv=None) -> DataFrame
    normalise_master(master_db) -> DataFrame          (indexed by join_key)
    fill_target(target_df, master_db, package_df, origin_df,
                options=FillOptions(), progress=None) -> (DataFrame, FillReport)
    run_ai_vision(target_df, image_dict, api_key, progress=None) -> AiVisionResult
    extract_images_from_zip(zip_file) -> {basename: bytes}
    write_filled_excel(df, path_or_buffer) -> None
    changed_columns(original_df, filled_df) -> [str]

The fill logic is a faithful extraction of the engine that previously lived
inline in app_manufacturer.py; only UI calls were replaced by the FillReport /
progress callback so behaviour is unchanged.
"""
from __future__ import annotations

import base64
import os
import re
import zipfile
from dataclasses import dataclass, field
from typing import Callable, Optional

import pandas as pd

from dictionaries import (
    TARGET_MAPPING,
    FACE_SHAPE_MAP,
    BRAND_USABLE_MAP,
    PREMIUM_KERING_BRANDS,
    BRAND_GLASSES_CONTAIN,
    estimate_filter_category,
)

# ==========================================================================
# Options / report containers
# ==========================================================================

@dataclass
class FillOptions:
    """Private-name numbers, keyed by glasses type. Empty string = don't set."""
    priv_sun: str = ""
    priv_eye: str = ""
    priv_pc: str = ""
    priv_sport: str = ""
    priv_drive: str = ""


@dataclass
class FillReport:
    total_rows: int = 0
    match_count: int = 0
    # column -> set of raw source values that contained "NOT MAPPED"
    unmapped: dict = field(default_factory=dict)
    # column -> count of rows whose source value was empty
    missing: dict = field(default_factory=dict)
    found_sport_glasses: bool = False
    found_polarized_clip_on: bool = False

    @property
    def unmatched_count(self) -> int:
        return max(0, self.total_rows - self.match_count)

    @property
    def total_issues(self) -> int:
        return sum(self.missing.values()) + sum(len(v) for v in self.unmapped.values())


@dataclass
class AiVisionResult:
    shape_count: int = 0
    sport_count: int = 0
    image_count: int = 0


ProgressFn = Optional[Callable[[float, str], None]]


def _tick(progress: ProgressFn, frac: float, text: str) -> None:
    if progress is not None:
        try:
            progress(max(0.0, min(1.0, frac)), text)
        except Exception:
            pass


# ==========================================================================
# File / data helpers
# ==========================================================================

def read_target_file(path_or_buffer, is_csv: Optional[bool] = None) -> pd.DataFrame:
    """Read a target import file and normalise its headers exactly as the web
    app does (line breaks -> space, collapse whitespace, strip)."""
    if is_csv is None:
        name = getattr(path_or_buffer, "name", str(path_or_buffer))
        is_csv = str(name).lower().endswith(".csv")

    if is_csv:
        df = pd.read_csv(path_or_buffer, dtype=str)
    else:
        df = pd.read_excel(path_or_buffer, dtype=str, engine="openpyxl")

    df.columns = (
        df.columns.astype(str)
        .str.replace("\n", " ", regex=False)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )
    return df


def normalise_master(master_db: pd.DataFrame) -> pd.DataFrame:
    """Ensure master_db is indexed by join_key (the web app does this at load
    time; the desktop app loads from a snapshot, so normalise here too)."""
    if master_db is None or master_db.empty:
        return pd.DataFrame()
    if master_db.index.name == "join_key":
        return master_db
    if "join_key" in master_db.columns:
        out = master_db.copy()
        out["join_key"] = out["join_key"].astype(str).str.strip()
        return out.set_index("join_key")
    return master_db


def target_barcode_column(target_df: pd.DataFrame) -> Optional[str]:
    """Return the barcode column name expected by TARGET_MAPPING, or None."""
    col = TARGET_MAPPING.get("Barcode", "Barcode")
    return col if col in target_df.columns else None


def changed_columns(original_df: pd.DataFrame, filled_df: pd.DataFrame) -> list:
    """Columns that the filler modified or added."""
    out = []
    for col in filled_df.columns:
        if col in original_df.columns:
            if not filled_df[col].equals(original_df[col]):
                out.append(col)
        else:
            out.append(col)
    return out


def write_filled_excel(df: pd.DataFrame, path_or_buffer, sheet_name: str = "Filled_Data") -> None:
    with pd.ExcelWriter(path_or_buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)


# ==========================================================================
# THE AUTO-FILLER ENGINE
# ==========================================================================

def fill_target(
    target_df: pd.DataFrame,
    master_db: pd.DataFrame,
    package_df: Optional[pd.DataFrame] = None,
    origin_df: Optional[pd.DataFrame] = None,
    options: FillOptions = None,
    progress: ProgressFn = None,
) -> tuple:
    """Fill a target import DataFrame from the master catalogue.

    Returns (filled_df, FillReport). The input DataFrame is copied, not mutated.
    """
    options = options or FillOptions()
    package_df = pd.DataFrame() if package_df is None else package_df
    origin_df = pd.DataFrame() if origin_df is None else origin_df

    target_df = target_df.copy()
    master_db = normalise_master(master_db)

    priv_sun = str(options.priv_sun or "").strip()
    priv_eye = str(options.priv_eye or "").strip()
    priv_pc = str(options.priv_pc or "").strip()
    priv_sport = str(options.priv_sport or "").strip()
    priv_drive = str(options.priv_drive or "").strip()

    target_barcode_col = TARGET_MAPPING.get("Barcode", "Barcode")
    if target_barcode_col not in target_df.columns:
        raise ValueError(f"Could not find the Barcode column '{target_barcode_col}' in the file.")

    report = FillReport(total_rows=len(target_df))

    # Ensure every mapped target column exists
    for global_col, target_col in TARGET_MAPPING.items():
        if isinstance(target_col, list):
            for tc in target_col:
                if tc not in target_df.columns:
                    target_df[tc] = ""
        else:
            if target_col not in target_df.columns:
                target_df[target_col] = ""

    unmapped_tracker = {}
    missing_tracker = {}

    # Caches for the majority engines
    brand_majority_cache = {}
    brand_contain_cache = {}
    brand_origin_cache = {}

    # Accept both "Glasses contain ID:84" (no space) and "Glasses contain ID: 84"
    # (with space) in templates — older templates used the space form, newer ones
    # don't. Whichever exists in the user's template is what we write to.
    if "Glasses contain ID:84" in target_df.columns:
        CONTAIN_COL = "Glasses contain ID:84"
    elif "Glasses contain ID: 84" in target_df.columns:
        CONTAIN_COL = "Glasses contain ID: 84"
    else:
        CONTAIN_COL = "Glasses contain ID:84"

    for c in ["Case length (mm)", "Case height (mm)", "Case width (mm)", "Case weight (g)", CONTAIN_COL]:
        if c not in target_df.columns:
            target_df[c] = ""

    total = len(target_df) or 1
    for step, (index, row) in enumerate(target_df.iterrows(), start=1):
        if step % 25 == 0 or step == total:
            _tick(progress, step / total, f"Matching barcodes… ({step}/{total})")

        raw_barcode = str(row[target_barcode_col]).strip()
        clean_barcode = re.sub(r"\.0$", "", raw_barcode).lstrip("0")

        if clean_barcode not in master_db.index:
            continue

        report.match_count += 1
        master_row = master_db.loc[clean_barcode]
        # A snapshot could in principle carry duplicate join_keys; take the first.
        if isinstance(master_row, pd.DataFrame):
            master_row = master_row.iloc[0]

        is_frames = str(master_row.get("Glasses_type", "")).strip() == "Frames"
        lens_cols_to_skip = [
            "Glasses_lens_Colour", "Glasses_lens_material",
            "Sunglasses_filter", "Glasses_lens_effect",
            "SunGlasses_RX_lenses",  # not applicable to optical frames
        ]

        target_df.at[index, "Items type ID: 20"] = "Glasses"
        target_df.at[index, "Items packing ID: 21"] = "Basic"

        g_type = str(master_row.get("Glasses_type", "")).strip()
        private_name = ""

        if "Sunglasses" in g_type and priv_sun:
            private_name = f"(Sunglasses {priv_sun})"
        elif "Sport glasses" in g_type and priv_sport:
            private_name = f"(Sports glasses {priv_sport})"
        elif "Driving glasses" in g_type and priv_drive:
            private_name = f"(Eyeglasses driving {priv_drive})"
        elif "PC Glasses without power" in g_type and priv_pc:
            private_name = f"(Eyeglasses PC {priv_pc})"
        elif "Frames" in g_type and priv_eye:
            private_name = f"(Eyeglasses {priv_eye})"

        if private_name:
            target_df.at[index, "Name private"] = private_name.strip()

        assembled_name = str(master_row.get("Assembled_Name", "")).strip()
        meta_desc = ""

        if "Sunglasses" in g_type:
            meta_desc = f"Sunglasses {assembled_name}"
        elif "Sport glasses" in g_type:
            meta_desc = f"Ski goggles {assembled_name}"
            report.found_sport_glasses = True
        elif "Driving glasses" in g_type:
            meta_desc = f"Driving glasses {assembled_name}"
        elif "PC Glasses without power" in g_type:
            meta_desc = f"Computer glasses {assembled_name}"
        elif "Frames" in g_type:
            meta_desc = f"Eyeglasses {assembled_name}"

        if meta_desc:
            target_df.at[index, "Meta description"] = meta_desc.strip()

        for global_col, target_col in TARGET_MAPPING.items():
            if global_col == "Barcode":
                continue
            if is_frames and global_col in lens_cols_to_skip:
                continue

            if global_col in master_db.columns:
                val = master_row[global_col]
                t_col_name = target_col[0] if isinstance(target_col, list) else target_col
                if pd.notna(val) and str(val).strip() != "":
                    val_str = str(val).strip()
                    # Filter out NOT MAPPED values (handles pipe-separated)
                    unmapped_parts = []
                    if "|" in val_str:
                        all_parts = [p.strip() for p in val_str.split("|") if p.strip()]
                        clean_parts = [p for p in all_parts if p != "NOT MAPPED"]
                        unmapped_parts = [p for p in all_parts if p == "NOT MAPPED"]
                        val_str = "|".join(clean_parts)
                    elif val_str == "NOT MAPPED":
                        unmapped_parts = [val_str]
                        val_str = ""
                    if unmapped_parts:
                        if t_col_name not in unmapped_tracker:
                            unmapped_tracker[t_col_name] = set()
                        unmapped_tracker[t_col_name].add(str(val).strip())
                    if val_str:
                        if isinstance(target_col, list):
                            for tc in target_col:
                                target_df.at[index, tc] = val_str
                        else:
                            target_df.at[index, target_col] = val_str
                    elif not val_str:
                        missing_tracker[t_col_name] = missing_tracker.get(t_col_name, 0) + 1
                else:
                    missing_tracker[t_col_name] = missing_tracker.get(t_col_name, 0) + 1

        g_shape_raw = str(master_row.get("Glasses_shape", "")).strip()
        if g_shape_raw and g_shape_raw.lower() not in ["nan", ""]:
            shapes = [s.strip() for s in g_shape_raw.split("|")]
            recommended_faces = set()
            for s in shapes:
                for shape_key, face_val in FACE_SHAPE_MAP.items():
                    if shape_key.lower() == s.lower():
                        for face in face_val.split("|"):
                            recommended_faces.add(face)
            if recommended_faces:
                target_df.at[index, "Glasses for your face shape ID:94"] = "|".join(sorted(list(recommended_faces)))

        if "Sunglasses" in g_type:
            target_df.at[index, "UV filter ID: 60"] = "400"

        # --- SUNGLASSES FILTER ESTIMATION (from lens colour) ---
        filter_col = "Sunglasses filter ID: 77"
        if filter_col in target_df.columns and "Sunglasses" in g_type:
            current_filter = str(target_df.at[index, filter_col]).strip()
            if not current_filter or current_filter.lower() in ["nan", ""]:
                raw_lens = str(master_row.get("Glasses_lens_Colour", "")).strip()
                estimated = estimate_filter_category(raw_lens)
                if estimated:
                    target_df.at[index, filter_col] = estimated

        usable_tags = set()
        raw_brand = str(master_row.get("Brand", "")).strip().lower()
        # Lookup candidates: each row's brand may have entries under the new
        # customer-facing long form ("Boss by Hugo Boss") OR the legacy short
        # form ("Hugo Boss") in any given reference table. Try long, then short.
        brand_lookup_candidates = [raw_brand]
        if raw_brand == "boss by hugo boss":
            brand_lookup_candidates.append("hugo boss")
        elif raw_brand == "hugo by hugo boss":
            brand_lookup_candidates.append("hugo")
        for _cand in brand_lookup_candidates:
            if _cand in BRAND_USABLE_MAP:
                usable_tags.add(BRAND_USABLE_MAP[_cand])
                break
        lens_effect = str(master_row.get("Glasses_lens_effect", "")).strip()

        if "Sunglasses" in g_type:
            if "Polarized" in lens_effect:
                usable_tags.add("Driving glasses")
            else:
                usable_tags.add("Common use")

        if usable_tags:
            target_df.at[index, "Glasses usable ID: 51"] = "|".join(sorted(list(usable_tags)))

        if any(c in PREMIUM_KERING_BRANDS for c in brand_lookup_candidates):
            target_df.at[index, "Glasses collection ID: 33"] = "Prémiové brýle - Kering"

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

        # --- CASE DIMENSIONS MAJORITY ENGINE ---
        if not package_df.empty and raw_brand and raw_brand != "nan":
            if raw_brand not in brand_majority_cache:
                brand_matches = pd.DataFrame()
                for _cand in brand_lookup_candidates:
                    mask = package_df["item_name"].astype(str).str.contains(
                        rf"\b{re.escape(_cand)}\b", case=False, na=False
                    )
                    if mask.any():
                        brand_matches = package_df[mask]
                        break

                if not brand_matches.empty:
                    def get_mode(col_name, _bm=brand_matches):
                        if col_name in _bm.columns:
                            modes = _bm[col_name].dropna().mode()
                            if not modes.empty:
                                return re.sub(r"\.0$", "", str(modes.iloc[0]).strip())
                        return ""

                    brand_majority_cache[raw_brand] = {
                        "Case length (mm)": get_mode("case_length"),
                        "Case height (mm)": get_mode("case_height"),
                        "Case width (mm)": get_mode("case_width"),
                        "Case weight (g)": get_mode("case_weight"),
                        "Glasses weight (g)": get_mode("item_weight"),
                    }
                else:
                    brand_majority_cache[raw_brand] = None

            cached_data = brand_majority_cache.get(raw_brand)
            if cached_data:
                target_df.at[index, "Case length (mm)"] = cached_data["Case length (mm)"]
                target_df.at[index, "Case height (mm)"] = cached_data["Case height (mm)"]
                target_df.at[index, "Case width (mm)"] = cached_data["Case width (mm)"]
                target_df.at[index, "Case weight (g)"] = cached_data["Case weight (g)"]
                if cached_data.get("Glasses weight (g)"):
                    target_df.at[index, "Glasses weight (g)"] = cached_data["Glasses weight (g)"]

        # --- ORIGIN COUNTRY MAJORITY ENGINE ---
        if not origin_df.empty and raw_brand and raw_brand != "nan":
            if raw_brand not in brand_origin_cache:
                if "item_name" in origin_df.columns and "country_master" in origin_df.columns:
                    brand_matches = pd.DataFrame()
                    for _cand in brand_lookup_candidates:
                        mask = origin_df["item_name"].astype(str).str.contains(
                            rf"\b{re.escape(_cand)}\b", case=False, na=False
                        )
                        if mask.any():
                            brand_matches = origin_df[mask]
                            break
                    if not brand_matches.empty:
                        modes = brand_matches["country_master"].dropna().mode()
                        brand_origin_cache[raw_brand] = str(modes.iloc[0]).strip() if not modes.empty else ""
                    else:
                        brand_origin_cache[raw_brand] = ""
                else:
                    brand_origin_cache[raw_brand] = ""

            cached_origin = brand_origin_cache.get(raw_brand, "")
            if cached_origin and "Item origin country" in target_df.columns:
                target_df.at[index, "Item origin country"] = cached_origin

        # --- GLASSES CONTAIN — STATIC BRAND × TYPE LOOKUP ---
        if raw_brand and raw_brand != "nan":
            if any(kw in g_type for kw in ("Sunglasses", "Sport glasses", "Driving glasses")):
                type_key = "Sunglasses"
            else:
                type_key = "Frames"
            contain_cache_key = (raw_brand, type_key)
            if contain_cache_key not in brand_contain_cache:
                resolved = ""
                for _cand in brand_lookup_candidates:
                    entry = BRAND_GLASSES_CONTAIN.get(_cand)
                    if entry and entry.get(type_key):
                        resolved = entry[type_key]
                        break
                brand_contain_cache[contain_cache_key] = resolved

            cached_contain = brand_contain_cache.get(contain_cache_key, "")

            clip_on_val = str(master_row.get("Extracted_Clip_on", "")).strip()
            needs_alert = master_row.get("Clip_on_Alert", False)

            if needs_alert:
                report.found_polarized_clip_on = True

            # Clip-on lens colour
            clip_lens_col = "Glasses clip-on lens colour ID:112"
            if clip_lens_col in target_df.columns:
                clip_lens_val = str(master_row.get("Clip_on_lens_colour", "")).strip()
                if clip_lens_val and clip_lens_val.lower() not in ["nan", ""]:
                    target_df.at[index, clip_lens_col] = clip_lens_val

            final_contain = []
            if cached_contain:
                final_contain.extend(cached_contain.split("|"))
            if clip_on_val and clip_on_val.lower() not in ["nan", ""]:
                final_contain.append(clip_on_val)

            if final_contain:
                unique_contain_dict = {item.strip().lower(): item.strip() for item in final_contain if item.strip()}
                ordered_items = []
                if "original glasses case" in unique_contain_dict:
                    ordered_items.append("Original glasses case")
                    del unique_contain_dict["original glasses case"]
                if "cleaning cloth" in unique_contain_dict:
                    ordered_items.append("Cleaning cloth")
                    del unique_contain_dict["cleaning cloth"]

                remaining_items = sorted(list(unique_contain_dict.values()))
                ordered_items.extend(remaining_items)
                target_df.at[index, CONTAIN_COL] = "|".join(ordered_items)

        # --- OTHER FEATURES ENGINE ---
        other_features = set()
        if "Glasses other features ID:99" in target_df.columns:
            existing_features = str(target_df.at[index, "Glasses other features ID:99"]).strip()
            if existing_features and existing_features.lower() not in ["nan", ""]:
                for e in existing_features.split("|"):
                    other_features.add(e.strip())

        # "Prescription sunglasses" only makes sense for actual sunglasses —
        # Safilo flags RXable=Y on optical frames too (any frame can be glazed),
        # so without this check the feature gets added to every row.
        if "Sunglasses" in g_type and "SunGlasses RX lenses ID:108" in target_df.columns:
            if str(target_df.at[index, "SunGlasses RX lenses ID:108"]).strip().lower() == "yes":
                other_features.add("Prescription sunglasses")

        if CONTAIN_COL in target_df.columns:
            contain_val = str(target_df.at[index, CONTAIN_COL]).strip().lower()
            contain_items = [item.strip() for item in re.split(r"[,|]", contain_val) if item.strip()]

            clip_on_found = False
            if "sun clip-on" in contain_items:
                other_features.add("Sun clip-on"); clip_on_found = True
            if "sun clip-on p" in contain_items:
                other_features.add("Sun clip-on p"); clip_on_found = True
            if "magnetic sun clip-on" in contain_items:
                other_features.add("Magnetic sun clip-on"); clip_on_found = True
            if "magnetic sun clip-on p" in contain_items:
                other_features.add("Magnetic sun clip-on p"); clip_on_found = True
            if clip_on_found:
                other_features.add("Glasses with sun clip-on")

        if other_features:
            target_df.at[index, "Glasses other features ID:99"] = "|".join(sorted(list(other_features)))

        # --- LENSES NO-ORDERS ENGINE ---
        no_orders = set()
        frame_type = (
            str(target_df.at[index, "Glasses frame type ID: 50"]).strip().lower()
            if "Glasses frame type ID: 50" in target_df.columns else ""
        )
        other_feat = (
            str(target_df.at[index, "Glasses other features ID:99"]).strip().lower()
            if "Glasses other features ID:99" in target_df.columns else ""
        )

        if frame_type == "half rim":
            no_orders.add("CoatingPolarized")
            no_orders.add("Glasses index 1.5")
        elif frame_type == "rimless":
            no_orders.add("CoatingPolarized")
            no_orders.add("Glasses index 1.5")
            no_orders.add("Glasses index 1.74")

        if "clip" in other_feat:
            no_orders.add("Glasses index 1.5")

        if no_orders and "Glasses lenses no-orders ID:103" in target_df.columns:
            target_df.at[index, "Glasses lenses no-orders ID:103"] = "|".join(sorted(no_orders))

        # --- Fill "None" for empty lens effect ---
        effect_col = "Glasses lens effect ID: 37"
        if effect_col in target_df.columns and not is_frames:
            current_effect = str(target_df.at[index, effect_col]).strip()
            if not current_effect or current_effect.lower() in ["nan", ""]:
                target_df.at[index, effect_col] = "None"

    report.unmapped = unmapped_tracker
    report.missing = missing_tracker
    _tick(progress, 1.0, f"Filled {report.match_count} of {report.total_rows} products.")
    return target_df, report


# ==========================================================================
# AI VISION (shape + sport detection) — optional
# ==========================================================================

SHAPE_CATEGORIES = [
    "Panthos / Tea cup", "Browline", "Cat Eye", "Oval / Elipse",
    "Butterfly", "Extravagant", "Single lens", "Square",
    "Oversize", "Hexagonal", "Pilot", "Rectangular", "Round",
]


def classify_glasses(image_bytes: bytes, api_key: str) -> dict:
    """Classify shape and sport type from a single image.
    Returns {'shape': str, 'is_sport': bool}; empty/False on any failure."""
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=api_key)

        img_b64 = base64.b64encode(image_bytes).decode("utf-8")
        ext_check = image_bytes[:8]
        media_type = "image/png" if ext_check[:4] == b"\x89PNG" else "image/jpeg"

        response = client.messages.create(
            model="claude-haiku-4-20250414",
            max_tokens=80,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": img_b64}},
                    {"type": "text", "text": (
                        "Analyze this eyewear image and answer TWO questions:\n\n"
                        "1. SHAPE: Classify into exactly ONE of these categories:\n"
                        + ", ".join(SHAPE_CATEGORIES) + "\n\n"
                        "2. SPORT: Are these sport/performance glasses (wrap-around, shield lens, "
                        "rubber grips, aerodynamic design, cycling/running/ski goggles)? Answer YES or NO.\n\n"
                        "Respond in exactly this format:\nSHAPE: <category>\nSPORT: <YES or NO>"
                    )},
                ],
            }],
        )
        result = response.content[0].text.strip()

        shape = ""
        is_sport = False
        for line in result.split("\n"):
            line = line.strip()
            if line.upper().startswith("SHAPE:"):
                raw_shape = line.split(":", 1)[1].strip()
                for cat in SHAPE_CATEGORIES:
                    if cat.lower() == raw_shape.lower():
                        shape = cat
                        break
                if not shape:
                    shape = raw_shape
            elif line.upper().startswith("SPORT:"):
                is_sport = "yes" in line.lower()

        return {"shape": shape, "is_sport": is_sport}
    except Exception:
        return {"shape": "", "is_sport": False}


def extract_images_from_zip(zip_file) -> dict:
    """{basename-without-extension: image bytes} from a ZIP path or buffer."""
    images = {}
    with zipfile.ZipFile(zip_file, "r") as z:
        for name in z.namelist():
            lower = name.lower()
            if lower.endswith((".jpg", ".jpeg", ".png")) and not name.startswith("__MACOSX"):
                basename = os.path.splitext(os.path.basename(name))[0]
                if basename:
                    images[basename] = z.read(name)
    return images


def load_images_from_folder(folder: str) -> dict:
    """Desktop affordance: read images straight from a folder (no ZIP needed)."""
    images = {}
    if not folder or not os.path.isdir(folder):
        return images
    for entry in os.listdir(folder):
        path = os.path.join(folder, entry)
        if not os.path.isfile(path):
            continue
        if entry.lower().endswith((".jpg", ".jpeg", ".png")):
            basename = os.path.splitext(entry)[0]
            if basename:
                try:
                    with open(path, "rb") as fh:
                        images[basename] = fh.read()
                except Exception:
                    pass
    return images


def run_ai_vision(
    target_df: pd.DataFrame,
    image_dict: dict,
    api_key: str,
    progress: ProgressFn = None,
) -> AiVisionResult:
    """Classify shapes from images and write them into target_df IN PLACE.
    Mirrors the web app: DB-sourced shapes are marked 'Database', AI ones 'AI'."""
    res = AiVisionResult(image_count=len(image_dict or {}))
    if not image_dict or not api_key:
        return res

    shape_col = "Glasses shape ID: 25"
    face_col = "Glasses for your face shape ID:94"
    sport_col = "Sports Glasses ID: 89"
    source_col = "Shape source"
    if shape_col not in target_df.columns:
        target_df[shape_col] = ""
    if face_col not in target_df.columns:
        target_df[face_col] = ""
    if sport_col not in target_df.columns:
        target_df[sport_col] = ""
    target_df[source_col] = ""

    # Mark existing shapes as coming from the database
    for idx, row in target_df.iterrows():
        if str(row.get(shape_col, "")).strip() not in ["", "nan"]:
            target_df.at[idx, source_col] = "Database"

    name_col = "Glasses name"
    if name_col not in target_df.columns:
        for c in target_df.columns:
            if "name" in c.lower() and "private" not in c.lower():
                name_col = c
                break

    total_rows = len(target_df) or 1
    for pos, (idx, row) in enumerate(target_df.iterrows(), start=1):
        glasses_name = str(row.get(name_col, "")).strip()
        if glasses_name and glasses_name in image_dict:
            result = classify_glasses(image_dict[glasses_name], api_key)

            if result["shape"]:
                target_df.at[idx, shape_col] = result["shape"]
                target_df.at[idx, source_col] = "AI"
                res.shape_count += 1

                recommended_faces = set()
                for shape_key, face_val in FACE_SHAPE_MAP.items():
                    if shape_key.lower() == result["shape"].lower():
                        for face in face_val.split("|"):
                            recommended_faces.add(face)
                if recommended_faces:
                    target_df.at[idx, face_col] = "|".join(sorted(recommended_faces))

            if result["is_sport"]:
                target_df.at[idx, sport_col] = "Yes"
                res.sport_count += 1

        _tick(progress, pos / total_rows, f"Classifying with AI vision… ({pos}/{total_rows})")

    return res
