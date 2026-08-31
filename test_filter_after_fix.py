# -*- coding: utf-8 -*-
import pandas as pd
import glob
from dictionaries import MANUFACTURER_CONFIG
from ingest import load_single_catalog

FILES = sorted(glob.glob(r"C:\Users\blank\Downloads\SafiloAvailability_CZ0120*.csv"))
print(f"Processing {len(FILES)} files...")

dfs = []
for f in FILES:
    cfg = MANUFACTURER_CONFIG["safilo"]
    df, _u, _s = load_single_catalog("safilo", cfg, f)
    dfs.append(df)
combined = pd.concat(dfs, ignore_index=True)
print(f"Combined: {len(combined):,} rows")
print()

print("Sunglasses_filter distribution AFTER fix:")
print(combined["Sunglasses_filter"].fillna("(empty)").value_counts(dropna=False).head(15).to_string())
print()
print("For comparison, the REAL distribution should be roughly:")
print("  Cat 3: ~44k, Cat 2: ~9k, Cat 1: ~1k, Cat 0: ~99, plus blanks for optical/special")
