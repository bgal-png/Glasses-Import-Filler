# -*- coding: utf-8 -*-
"""Verify the new clip-on detection logic against all 11 files."""
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

# Replicate the new clip-on engine
clip_vals = []
for _, row in df.iterrows():
    style_d = str(row.get("StyleD", "")).strip().upper()
    pol = str(row.get("LenPolarized", "")).strip().upper()
    prod_type = re.sub(r"\s+", " ", str(row.get("TypeD", "")).strip().upper())

    is_clipon = (
        style_d.endswith("/C")
        or "CLIP-IN" in style_d
        or "CLIP-ON" in style_d
        or "CLIP ON" in style_d
        or "CLIP-ON" in prod_type
        or "CLIP ON" in prod_type
    )
    if is_clipon:
        clip_vals.append("Magnetic sun clip-on p" if pol in ("X", "Y") else "Magnetic sun clip-on")
    else:
        clip_vals.append("")

df["Extracted_Clip_on"] = clip_vals

total = len(df)
clipons = df[df["Extracted_Clip_on"].ne("")]
print(f"Total rows: {total:,}")
print(f"Rows flagged as clip-on: {len(clipons):,}  ({len(clipons)/total*100:.1f}%)")
print()
print("Breakdown by extracted value:")
print(clipons["Extracted_Clip_on"].value_counts().to_string())
print()
print("Breakdown by detection trigger:")
print(f"  StyleD ends in /C : {df['StyleD'].astype(str).str.upper().str.endswith('/C').sum():,}")
print(f"  StyleD has CLIP-IN: {df['StyleD'].astype(str).str.upper().str.contains('CLIP-IN', na=False).sum():,}")
print(f"  StyleD has CLIP-ON: {df['StyleD'].astype(str).str.upper().str.contains('CLIP-ON', na=False).sum():,}")
print()
print("Sample clip-on rows (first 15):")
print(clipons[["BrandD", "StyleD", "TypeD", "LenPolarized", "LenColorD", "Extracted_Clip_on"]].head(15).to_string())
print()

# Sanity check: are any clip-on rows NOT optical frames?
print("Clip-on rows by TypeD:")
print(clipons["TypeD"].value_counts().to_string())
print()

# Sanity check: how many clip-ons have polarized=N (non-polarized)?
print("Clip-on rows by LenPolarized:")
print(clipons["LenPolarized"].value_counts().to_string())
print()

# Cross-check: any /C StyleD that we incorrectly tagged?
print("Distinct StyleD endings flagged as clip-on (sample):")
distinct_styles = sorted(set(clipons["StyleD"].astype(str)))[:20]
for s in distinct_styles:
    print(f"  '{s}'")
print(f"  ... ({len(set(clipons['StyleD']))} distinct StyleD values total)")
