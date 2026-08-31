# -*- coding: utf-8 -*-
"""Hunt for clip-on signals across every column in every file."""
import pandas as pd
import glob
import re

FILES = sorted(glob.glob(r"C:\Users\blank\Downloads\SafiloAvailability_CZ0120*.csv"))

all_dfs = []
for f in FILES:
    d = pd.read_csv(f, dtype=str, on_bad_lines="skip", sep=";")
    if len(d.columns) <= 1:
        d = pd.read_csv(f, dtype=str, on_bad_lines="skip", sep=",")
    d.columns = d.columns.astype(str).str.strip()
    all_dfs.append(d)

df = pd.concat(all_dfs, ignore_index=True)
print(f"Loaded {len(df):,} rows from {len(FILES)} files")
print(f"Columns: {list(df.columns)}")
print()

# --- Search every text column for "clip" (case-insensitive) ---
print("=" * 60)
print("SEARCH: any column containing 'clip' (case-insensitive)")
print("=" * 60)
for col in df.columns:
    if df[col].dtype != "object": continue
    matches = df[col].astype(str).str.contains(r"clip", case=False, na=False, regex=True)
    n = matches.sum()
    if n > 0:
        print(f"\n[{col}]  {n:,} matches")
        # Show unique values containing 'clip'
        vals = df.loc[matches, col].astype(str).str.strip().unique()
        for v in sorted(set(vals))[:20]:
            count = (df[col].astype(str).str.strip() == v).sum()
            print(f"    {count:>6,}x  '{v}'")

print()
print("=" * 60)
print("EXAMINE: every column's distinct values for promising clip-related fields")
print("=" * 60)

# Fields that might hold clip-on info
candidates = ["Fitting", "Type", "TypeD", "Style", "StyleD", "Case", "CaseD",
              "CaseGroup", "CoreCollection", "Shape", "BrandD"]
for col in candidates:
    if col not in df.columns: continue
    uniques = sorted(set(str(x).strip() for x in df[col].dropna().unique() if str(x).strip()))
    print(f"\n[{col}]  {len(uniques)} unique values:")
    if len(uniques) <= 50:
        for u in uniques:
            count = (df[col].astype(str).str.strip() == u).sum()
            print(f"    {count:>6,}x  '{u}'")
    else:
        # Just show any that look clip-related
        clipish = [u for u in uniques if "clip" in u.lower() or "+" in u]
        print(f"    (too many to list; clip/+ candidates: {clipish[:20]})")

# --- Look at what Fitting column contains specifically ---
if "Fitting" in df.columns:
    print()
    print("=" * 60)
    print("Fitting column full distribution")
    print("=" * 60)
    counts = df["Fitting"].value_counts(dropna=False).head(30)
    print(counts.to_string())

# --- Sample rows where Polarized=Y and TypeD=Optical frames (potential clip-on bundles) ---
print()
print("=" * 60)
print("Optical frames with Polarized=Y (suspicious — usually means clip-on)")
print("=" * 60)
if "TypeD" in df.columns and "LenPolarized" in df.columns:
    mask = (df["TypeD"].astype(str).str.strip().eq("Optical frames")) & \
           (df["LenPolarized"].astype(str).str.strip().eq("Y"))
    print(f"Count: {mask.sum():,}")
    if mask.sum():
        print(df.loc[mask, ["BrandD", "StyleD", "Style", "Upc", "TypeD", "LenPolarized",
                            "LenColorD", "Fitting"]].head(10).to_string())
