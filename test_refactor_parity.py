# -*- coding: utf-8 -*-
"""Sanity check: confirm ingest.py produces a sensible DataFrame on a real file
and that the public API of ingest.py matches what app_admin and the future
headless script need."""
import pandas as pd
from dictionaries import MANUFACTURER_CONFIG
from ingest import load_single_catalog, perform_upsert

FILE = r"C:\Users\blank\Downloads\SafiloAvailability_CZ0120 (1).csv"

cfg = MANUFACTURER_CONFIG["safilo"]
df, unmapped, skipped = load_single_catalog("safilo", cfg, FILE)

print(f"Returned columns: {len(df.columns)}")
print(f"Returned rows:    {len(df):,}")
print(f"Unmapped values:  {len(unmapped)}")
print(f"Skipped NM:       {len(skipped)}")
print()
print("Expected key columns present:")
for c in ["join_key", "Barcode", "Brand", "Manufacturer", "Combination",
          "Extracted_Clip_on", "Producing_company", "Assembled_Name"]:
    print(f"  {c}: {'OK' if c in df.columns else 'MISSING'}")
print()
print("Sample 3 rows (selected cols):")
cols = ["join_key", "Brand", "Combination", "Glasses_type", "Glasses_main_material",
        "Extracted_Clip_on", "SunGlasses_RX_lenses"]
present_cols = [c for c in cols if c in df.columns]
print(df[present_cols].head(3).to_string())
print()
print("OK." if not unmapped and not skipped else f"Warnings: {len(unmapped) + len(skipped)}")
