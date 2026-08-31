# -*- coding: utf-8 -*-
"""
Pure (UI-free) admin operations: catalogue ingest, renaming, the created-items
registry and the danger-zone delete.

No Streamlit, no Qt — the desktop admin tabs and (optionally) the Streamlit
admin app both drive these. Every write goes through here so the behaviour
can't drift between the two front ends.
"""
from __future__ import annotations

import re
from typing import Callable, Optional

import pandas as pd

from dictionaries import MANUFACTURER_CONFIG
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
# Single-row lookup and edit
# ==========================================================================

def fetch_row(engine, barcode: str):
    """One master_catalog row as a Series, or None. Targeted query — never
    pulls the whole table."""
    from sqlalchemy import text

    key = clean_barcode(barcode)
    if not key or key == "nan":
        return None
    df = pd.read_sql(
        text('SELECT * FROM master_catalog WHERE join_key = :k LIMIT 1'),
        con=engine, params={"k": key},
    )
    return None if df.empty else df.iloc[0]


def update_row(engine, barcode: str, changes: dict, progress: ProgressFn = None) -> dict:
    """Write only the changed cells of one row.

    A targeted UPDATE, deliberately: the Streamlit version reads the entire
    table and rewrites it with to_sql(if_exists="replace") to change a single
    cell, which rewrites 100k+ rows and briefly drops the table.

    Column names are whitelisted against the live table before being
    interpolated, since an identifier can't be a bound parameter.
    """
    from sqlalchemy import text

    key = clean_barcode(barcode)
    if not key:
        raise ValueError("No barcode given.")
    changes = {k: v for k, v in (changes or {}).items() if k != "join_key"}
    if not changes:
        return {"updated": 0, "columns": []}

    _tick(progress, 0.2, "Checking the row…")
    existing = pd.read_sql(
        text('SELECT * FROM master_catalog WHERE join_key = :k LIMIT 1'),
        con=engine, params={"k": key},
    )
    if existing.empty:
        raise ValueError(f"Barcode {barcode} is not in the catalogue.")

    valid = set(existing.columns)
    unknown = [c for c in changes if c not in valid]
    if unknown:
        raise ValueError(f"Unknown column(s): {', '.join(sorted(unknown))}")

    assignments = ", ".join(f'"{c}" = :v{i}' for i, c in enumerate(changes))
    params = {f"v{i}": ("" if v is None else str(v)) for i, v in enumerate(changes.values())}
    params["k"] = key

    _tick(progress, 0.6, f"Updating {len(changes)} field(s)…")
    with engine.begin() as conn:
        conn.execute(
            text(f'UPDATE master_catalog SET {assignments} WHERE join_key = :k'),
            params,
        )
    _tick(progress, 1.0, "Done.")
    return {"updated": len(changes), "columns": sorted(changes)}


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
