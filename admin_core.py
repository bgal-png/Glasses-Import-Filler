# -*- coding: utf-8 -*-
"""
Pure (UI-free) admin operations: catalogue ingest, colour filling, renaming,
the created-items registry and the danger-zone delete.

No Streamlit, no Qt — the desktop admin tabs and (optionally) the Streamlit
admin app both drive these. Every write goes through here so the behaviour
can't drift between the two front ends.
"""
from __future__ import annotations

import re
from typing import Callable, Optional

import pandas as pd

from dictionaries import BRAND_GLASSES_CONTAIN, MANUFACTURER_CONFIG  # noqa: F401
from ingest import load_single_catalog, perform_upsert, record_ingest

ProgressFn = Optional[Callable[[float, str], None]]


def _tick(progress: ProgressFn, frac: float, text: str) -> None:
    if progress is not None:
        try:
            progress(max(0.0, min(1.0, frac)), text)
        except Exception:
            pass


_ENGINES: dict = {}


def make_engine(db_url: str):
    from sqlalchemy import create_engine
    return create_engine(db_url, pool_pre_ping=True, pool_recycle=300)


def get_engine(db_url: str):
    """Cached engine per URL — creating one per operation would leak pools."""
    if not db_url:
        raise ValueError(
            "No database URL configured. Paste it in Settings to unlock admin actions."
        )
    if db_url not in _ENGINES:
        _ENGINES[db_url] = make_engine(db_url)
    return _ENGINES[db_url]


def clean_barcode(x) -> str:
    """Normalise a barcode the way master_catalog stores join_key."""
    return re.sub(r"\.0$", "", str(x).strip()).lstrip("0")


# ==========================================================================
# Catalogue ingest
# ==========================================================================

def process_catalogue(engine, mfg: str, file_path: str, progress: ProgressFn = None) -> dict:
    """Run a manufacturer file through the rules engine and upsert it.

    Mirrors the admin app exactly, including the brand-expansion step (the
    upsert dedupes by join_key afterwards).
    """
    if mfg not in MANUFACTURER_CONFIG:
        raise ValueError(f"Unknown manufacturer '{mfg}'.")
    config = MANUFACTURER_CONFIG[mfg]

    _tick(progress, 0.1, f"Reading and translating the {mfg.title()} file…")
    df, unmapped, skipped = load_single_catalog(mfg, config, file_path)
    if df.empty:
        return {
            "rows": 0, "unique": 0, "unmapped": sorted(unmapped),
            "skipped": sorted(skipped), "message": "No rows could be extracted.",
        }

    unique = df["join_key"].nunique() if "join_key" in df.columns else len(df)

    _tick(progress, 0.6, "Expanding by brand…")
    expanded = pd.concat([df.copy() for _ in config["brands"]], ignore_index=True) \
        if config.get("brands") else df

    _tick(progress, 0.75, "Upserting into master_catalog…")
    message = perform_upsert(expanded, engine)

    _tick(progress, 0.95, "Recording the ingest timestamp…")
    record_ingest(engine, mfg, unique)

    _tick(progress, 1.0, "Done.")
    return {
        "rows": int(len(df)), "unique": int(unique),
        "unmapped": sorted(unmapped), "skipped": sorted(skipped),
        "message": message,
    }


def delete_manufacturer(engine, mfg: str, progress: ProgressFn = None) -> dict:
    """Remove every master_catalog row whose Producing_company matches."""
    _tick(progress, 0.2, "Loading master_catalog…")
    full = pd.read_sql_table("master_catalog", con=engine)
    before = len(full)
    if "Producing_company" not in full.columns:
        raise ValueError("master_catalog has no 'Producing_company' column.")

    target = mfg.title().lower()
    keep = full[full["Producing_company"].astype(str).str.strip().str.lower() != target]
    deleted = before - len(keep)
    if deleted:
        _tick(progress, 0.7, f"Writing back {len(keep):,} remaining rows…")
        keep.to_sql("master_catalog", con=engine, if_exists="replace", index=False)
    _tick(progress, 1.0, "Done.")
    return {"deleted": int(deleted), "remaining": int(len(keep))}


# ==========================================================================
# Rename by barcode
# ==========================================================================

def apply_renames(engine, name_by_barcode: dict, progress: ProgressFn = None) -> dict:
    """Set Assembled_Name for every matching barcode. Keys may be raw barcodes."""
    mapping = {}
    for raw, name in name_by_barcode.items():
        key = clean_barcode(raw)
        nm = str(name).strip()
        if key and key != "nan" and nm and nm.lower() != "nan":
            mapping[key] = nm
    if not mapping:
        return {"updated": 0, "not_found": []}

    _tick(progress, 0.2, "Loading master_catalog…")
    full = pd.read_sql_table("master_catalog", con=engine)
    full["join_key"] = full["join_key"].astype(str).str.strip()
    if "Assembled_Name" not in full.columns:
        full["Assembled_Name"] = ""

    mask = full["join_key"].isin(mapping.keys())
    _tick(progress, 0.6, f"Renaming {int(mask.sum())} row(s)…")
    for pos in full.index[mask]:
        full.at[pos, "Assembled_Name"] = mapping[full.at[pos, "join_key"]]

    found = set(full.loc[mask, "join_key"])
    not_found = sorted(k for k in mapping if k not in found)

    _tick(progress, 0.85, "Writing back…")
    full.to_sql("master_catalog", con=engine, if_exists="replace", index=False)
    _tick(progress, 1.0, "Done.")
    return {"updated": int(mask.sum()), "not_found": not_found}


# ==========================================================================
# Colours from photos
# ==========================================================================

COLOUR_FIELDS = [
    # (master_catalog column, label, palette, applies-when)
    ("Frame_Colour", "Frame colour", "frame", "any"),
    ("Temple_Colour", "Temple colour", "frame", "any"),
    ("Glasses_lens_Colour", "Lens colour", "lens", "sunglasses"),
    ("Clip_on_lens_colour", "Clip-on lens colour", "lens", "clip"),
]
GRADIENT_MARKER = "__gradient__"


def _blank(v) -> bool:
    return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() in ("", "nan")


def norm_key(s) -> str:
    """Alphanumeric-only uppercase, for matching filenames against codes."""
    return re.sub(r"[^A-Za-z0-9]", "", str(s or "")).upper()


def build_colour_worklist(master_db: pd.DataFrame, producers: list | None = None) -> list:
    """Group rows by (model, colour code) and list which colour fields are
    missing. Colour is shared across sizes, so one decision fills every
    barcode in the group. Pure — no DB, no photos.
    """
    if master_db is None or master_db.empty:
        return []
    df = master_db
    if df.index.name == "join_key":
        df = df.reset_index()
    if producers and "Producing_company" in df.columns:
        wanted = {str(p).strip().lower() for p in producers}
        df = df[df["Producing_company"].astype(str).str.strip().str.lower().isin(wanted)]

    groups = {}
    for _, row in df.iterrows():
        model = str(row.get("Extracted_Model", "")).strip()
        colour = str(row.get("Glasses_color_code", "")).strip()
        if not model or model.lower() == "nan":
            continue
        g_type = str(row.get("Glasses_type", "")).strip()
        has_clip = not _blank(row.get("Extracted_Clip_on", ""))

        missing = []
        for col, _label, _palette, cond in COLOUR_FIELDS:
            if not _blank(row.get(col, "")):
                continue
            if cond == "sunglasses" and "Sunglasses" not in g_type:
                continue
            if cond == "clip" and not has_clip:
                continue
            missing.append(col)

        lens_effect = str(row.get("Glasses_lens_effect", "")).strip()
        can_gradient = ("Sunglasses" in g_type) and ("gradient" not in lens_effect.lower())
        if not missing and not can_gradient:
            continue

        key = (norm_key(model), norm_key(colour))
        if key not in groups:
            groups[key] = {
                "key": f"{key[0]}|{key[1]}",
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

    out = []
    for (model_n, colour_n), g in groups.items():
        g["barcodes"] = sorted(g["barcodes"])
        g["missing"] = [c for c, _l, _p, _cd in COLOUR_FIELDS if c in g["missing"]]
        g["_model_n"] = model_n
        g["_colour_n"] = colour_n
        out.append(g)
    out.sort(key=lambda g: (g["brand"], g["model"], g["colour_code"]))
    return out


def match_photos(worklist: list, photo_names: dict) -> tuple:
    """Attach a photo to each group.

    `photo_names` maps an identifier (path or zip member) to its basename.
    Matching is on model + colour code found anywhere in the normalised
    filename, and also accepts the size+colour form because the stored colour
    code has the size stripped off.

    Returns (matched_groups, unmatched_groups).
    """
    index = [(norm_key(base), ident) for ident, base in photo_names.items()]
    matched, unmatched = [], []
    for g in worklist:
        model_n = g["_model_n"]
        colour_n = g["_colour_n"]
        sizecolour_n = norm_key(g.get("size", "")) + colour_n
        found = None
        for norm_name, ident in index:
            if not model_n or model_n not in norm_name:
                continue
            if (colour_n and colour_n in norm_name) or (sizecolour_n and sizecolour_n in norm_name):
                found = ident
                break
        if found:
            item = dict(g)
            item["photo"] = found
            matched.append(item)
        else:
            unmatched.append(g)
    return matched, unmatched


def save_colours(engine, assignments: dict, barcodes_by_group: dict,
                 progress: ProgressFn = None) -> dict:
    """Write chosen colours to every barcode in each group.

    `assignments`: {group_key: {master_column: value, '__gradient__': True}}
    """
    if not assignments:
        return {"cells": 0, "groups": 0}

    _tick(progress, 0.2, "Loading master_catalog…")
    full = pd.read_sql_table("master_catalog", con=engine)
    full["join_key"] = full["join_key"].astype(str).str.strip()

    def add_gradient(v):
        parts = [p.strip() for p in str(v or "").split("|")
                 if p.strip() and p.strip().lower() != "nan"]
        if "Gradient" not in parts:
            parts.append("Gradient")
        return "|".join(sorted(set(parts)))

    cells = 0
    for gkey, fields in assignments.items():
        barcodes = set(barcodes_by_group.get(gkey, []))
        if not barcodes:
            continue
        mask = full["join_key"].isin(barcodes)
        for field, value in fields.items():
            if field == GRADIENT_MARKER:
                if "Glasses_lens_effect" not in full.columns:
                    full["Glasses_lens_effect"] = ""
                full.loc[mask, "Glasses_lens_effect"] = \
                    full.loc[mask, "Glasses_lens_effect"].apply(add_gradient)
            else:
                if field not in full.columns:
                    full[field] = ""
                full.loc[mask, field] = value
            cells += int(mask.sum())

    _tick(progress, 0.8, "Writing back…")
    full.to_sql("master_catalog", con=engine, if_exists="replace", index=False)
    _tick(progress, 1.0, "Done.")
    return {"cells": cells, "groups": len(assignments)}


# ==========================================================================
# Created-items registry
# ==========================================================================

REGISTRY_TABLE = "created_items"


def store_created_items(engine, records: pd.DataFrame, progress: ProgressFn = None) -> dict:
    """Merge (join_key, barcode, name, size) rows into the registry."""
    if records is None or records.empty:
        return {"incoming": 0, "total": 0, "added": 0}

    _tick(progress, 0.3, "Loading the registry…")
    try:
        existing = pd.read_sql_table(REGISTRY_TABLE, con=engine)
    except Exception:
        existing = pd.DataFrame(columns=["join_key", "barcode", "name", "size"])
    before = len(existing)

    combined = pd.concat([existing, records], ignore_index=True)
    combined.drop_duplicates(subset=["join_key"], keep="last", inplace=True)

    _tick(progress, 0.8, "Writing back…")
    combined.to_sql(REGISTRY_TABLE, con=engine, if_exists="replace", index=False)
    _tick(progress, 1.0, "Done.")
    return {
        "incoming": int(len(records)),
        "total": int(len(combined)),
        "added": int(len(combined) - before),
    }


def check_created_items(engine, barcodes: list) -> pd.DataFrame:
    """Which of these barcodes have we already created? -> Barcode/Status/Name/Size."""
    try:
        registry = pd.read_sql_table(REGISTRY_TABLE, con=engine)
    except Exception:
        registry = pd.DataFrame(columns=["join_key", "barcode", "name", "size"])
    if registry.empty:
        raise ValueError("The registry is empty — store some filled files first.")

    registry["join_key"] = registry["join_key"].astype(str).str.strip()
    lookup = registry.drop_duplicates(subset=["join_key"], keep="last").set_index("join_key")

    rows = []
    for raw in barcodes:
        raw_s = str(raw).strip()
        if not raw_s or raw_s.lower() == "nan":
            continue
        key = clean_barcode(raw_s)
        if key in lookup.index:
            r = lookup.loc[key]
            rows.append({
                "Barcode": raw_s, "Status": "Already created",
                "Name": str(r.get("name", "") or ""), "Size": str(r.get("size", "") or ""),
            })
        else:
            rows.append({"Barcode": raw_s, "Status": "New", "Name": "", "Size": ""})
    return pd.DataFrame(rows)


def records_from_filled_file(df: pd.DataFrame) -> pd.DataFrame:
    """Pull (join_key, barcode, name, size) out of a filled import file."""
    def find(cands):
        lower = {str(c).lower(): c for c in df.columns}
        for cand in cands:
            if cand.lower() in lower:
                return lower[cand.lower()]
        return None

    bc = find(["Barcode", "EAN", "UPC", "EAN/UPC"])
    if bc is None:
        raise ValueError("No Barcode column found.")
    name = find(["Glasses name", "XML description", "Name"])
    size = find(["Combination (size on glasses)", "Combination", "Size"])

    out = pd.DataFrame()
    out["join_key"] = df[bc].apply(clean_barcode)
    out["barcode"] = df[bc].astype(str).str.strip()
    out["name"] = df[name].astype(str).str.strip() if name else ""
    out["size"] = df[size].astype(str).str.strip() if size else ""
    return out[(out["join_key"] != "") & (out["join_key"] != "nan")]
