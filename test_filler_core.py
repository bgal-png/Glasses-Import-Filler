# -*- coding: utf-8 -*-
"""End-to-end smoke test for the extracted filler_core, with NO database.

Builds a real master_db locally by running the ingest engine over a catalogue
file, then fills a real target template and reports what happened.
"""
import sys

import pandas as pd

from dictionaries import MANUFACTURER_CONFIG
from ingest import load_single_catalog
from filler_core import (
    FillOptions, changed_columns, fill_target, read_target_file,
)

CATALOGUE = r"C:\Users\blank\Downloads\SafiloAvailability_CZ0120 (11).csv"
TARGET = r"C:\Users\blank\Desktop\OLD PC\BG Old\Glasses Imports\filler-testing.xlsx"

print("Building master_db from catalogue (no DB)…")
mdf, unmapped, _ = load_single_catalog("safilo", MANUFACTURER_CONFIG["safilo"], CATALOGUE)
print(f"  master_db: {len(mdf):,} rows, {len(unmapped)} unmapped")

print("Reading target template (for its real 51 headers)…")
template = read_target_file(TARGET)
print(f"  template: {len(template)} rows x {len(template.columns)} cols")

# The template's own barcode is a Luxottica one, so build a target that uses the
# template's REAL headers with real barcodes from this catalogue — a mix of
# sunglasses, optical frames and clip-ons so every branch of the engine runs.
sample_keys = []
for gtype in ("Sunglasses", "Frames"):
    sub = mdf[mdf["Glasses_type"] == gtype]
    sample_keys += list(sub["join_key"].head(6))
clip = mdf[mdf["Extracted_Clip_on"].astype(str).str.strip() != ""]
sample_keys += list(clip["join_key"].head(3))
raw_barcodes = list(mdf.set_index("join_key").loc[sample_keys, "Barcode"])

target = pd.DataFrame(columns=template.columns)
target["Barcode"] = raw_barcodes
target = target.astype(object)
original = target.copy()
print(f"  target: {len(target)} rows x {len(target.columns)} cols "
      f"(6 sunglasses + 6 frames + 3 clip-ons)")

ticks = []
filled, report = fill_target(
    target, mdf,
    package_df=pd.DataFrame(), origin_df=pd.DataFrame(),
    options=FillOptions(priv_sun="1001", priv_eye="2001"),
    progress=lambda f, t: ticks.append((round(f, 2), t)),
)

print()
print("=== REPORT ===")
print(f"  total_rows            : {report.total_rows}")
print(f"  match_count           : {report.match_count}")
print(f"  unmatched             : {report.unmatched_count}")
print(f"  unmapped cols         : {len(report.unmapped)}")
print(f"  missing cols          : {len(report.missing)}")
print(f"  found_sport_glasses   : {report.found_sport_glasses}")
print(f"  found_polarized_clip  : {report.found_polarized_clip_on}")
print(f"  total_issues          : {report.total_issues}")
print(f"  progress callbacks    : {len(ticks)} (last={ticks[-1] if ticks else None})")

cols = changed_columns(original, filled)
print()
print(f"=== {len(cols)} COLUMNS CHANGED ===")
for c in cols:
    print(f"  {c}")

key_cols = [c for c in [
    "Barcode", "Glasses name", "Meta description", "Item description", "HS Code",
    "Items type ID: 20", "Glasses type ID: 13", "Manufacturer ID: 9", "Brand ID:11",
    "Glasses usable ID: 51", "Glasses for your face shape ID:94", "Name private",
    "Combination (size on glasses)", "Glasses color code ID:107",
] if c in filled.columns]
print()
print("=== SAMPLE OF FILLED ROWS ===")
print(filled[key_cols].head(6).to_string())

assert report.match_count > 0, "FAIL: nothing matched — extraction is broken"
assert len(cols) > 5, "FAIL: suspiciously few columns changed"
print()
print("PASS: core filled data end-to-end with no UI framework loaded.")
print("Streamlit imported?", "streamlit" in sys.modules)
print("Qt imported?", any(m.startswith("PySide") for m in sys.modules))
