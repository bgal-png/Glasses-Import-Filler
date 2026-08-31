# -*- coding: utf-8 -*-
"""Dry-run test across multiple Safilo CSVs. Combines all files and reports
column mapping, combination construction, clip-on detection, unmapped values.
"""
import pandas as pd
import re
import glob
from dictionaries import MANUFACTURER_CONFIG, VALUE_TRANSLATOR, KNOWN_BRANDS, classify_color

FILES = sorted(glob.glob(r"C:\Users\blank\Downloads\SafiloAvailability_CZ0120*.csv"))
MFG = "safilo"
config = MANUFACTURER_CONFIG[MFG]

print(f"=== FOUND {len(FILES)} FILES ===")
for f in FILES:
    print(f"  - {f.rsplit(chr(92), 1)[-1]}")
print()

all_dfs = []
for f in FILES:
    d = pd.read_csv(f, dtype=str, on_bad_lines="skip", sep=",")
    if len(d.columns) <= 1:
        d = pd.read_csv(f, dtype=str, on_bad_lines="skip", sep=";")
    d.columns = d.columns.astype(str).str.strip()
    d["__source_file"] = f.rsplit(chr(92), 1)[-1]
    all_dfs.append(d)

df = pd.concat(all_dfs, ignore_index=True)
print(f"=== COMBINED TOTAL ROWS: {len(df):,} ===")
print(f"Columns: {len(df.columns)}")
print()

# --- CHECK EXPECTED COLUMNS EXIST ---
expected = set()
for v in config["columns"].values():
    if isinstance(v, str) and v:
        expected.add(v)
    elif isinstance(v, list):
        expected.update(v)
missing = expected - set(df.columns)
present = expected & set(df.columns)
print(f"=== COLUMN MAPPING CHECK ===")
print(f"Mapped columns FOUND: {len(present)}/{len(expected)}")
if missing:
    print(f"!!! MISSING: {sorted(missing)}")
else:
    print("All expected source columns present.")
print()

# --- CLIP-ON DETECTION ---
clip_count = 0
clip_examples = []
for _, raw_row in df.iterrows():
    prod_type = re.sub(r"\s+", " ", str(raw_row.get("TypeD", "")).strip().upper())
    if "CLIP-ON" in prod_type or "CLIP ON" in prod_type:
        clip_count += 1
        if len(clip_examples) < 5:
            clip_examples.append({
                "TypeD": raw_row.get("TypeD"),
                "LenPolarized": raw_row.get("LenPolarized"),
                "StyleD": raw_row.get("StyleD"),
                "Upc": raw_row.get("Upc"),
            })
print(f"=== CLIP-ON DETECTION ===")
print(f"Rows with CLIP-ON in TypeD: {clip_count}")
for e in clip_examples:
    print(f"  {e}")
print()

# --- UNIQUE VALUES IN KEY FIELDS ---
print(f"=== UNIQUE VALUE SAMPLES (all files combined) ===")
for col in ["TypeD", "Material", "Rim", "Shape", "Gender", "Hinge",
            "LenPolarized", "LenPhotochromic", "RXable", "LenMaterial",
            "LenFilterCategory"]:
    if col in df.columns:
        uniques = sorted(set(str(x).strip() for x in df[col].dropna().unique() if str(x).strip()))
        print(f"  {col} ({len(uniques)} unique): {uniques[:30]}")
print()

# --- TRANSLATION CHECK ---
print(f"=== TRANSLATION CHECK ===")
for global_col, src_col in [
    ("Glasses_type", "TypeD"),
    ("Glasses_main_material", "Material"),
    ("Glasses_frame_type", "Rim"),
    ("Glasses_shape", "Shape"),
    ("Glasses_gendre", "Gender"),
    ("Glasses_lens_material", "LenMaterial"),
]:
    if src_col not in df.columns: continue
    translator = VALUE_TRANSLATOR.get(global_col, {})
    lower = {str(k).lower(): v for k, v in translator.items() if k}
    unmapped = set()
    mapped = set()
    for val in df[src_col].dropna().unique():
        v = str(val).strip()
        if not v: continue
        for part in [p.strip() for p in v.split(",") if p.strip()]:
            if part.lower() in lower:
                mapped.add(part)
            else:
                unmapped.add(part)
    print(f"  [{global_col}]  mapped:{len(mapped)}  unmapped:{len(unmapped)}")
    if unmapped:
        print(f"     !!! UNMAPPED: {sorted(unmapped)}")
print()

# --- BRAND WHITELIST CHECK ---
print(f"=== BRAND WHITELIST CHECK ===")
brand_lookup = sorted(KNOWN_BRANDS, key=len, reverse=True)
def clean_brand(raw):
    raw = str(raw).strip()
    if not raw or raw.lower() == "nan": return raw
    rl = raw.lower()
    for known in brand_lookup:
        if rl == known.lower(): return known
        if rl.startswith(known.lower() + " "): return known
    return raw

BRAND_CORRECTIONS = {
    "moschino love": "Love Moschino",
    "prive' revaux": "Prive Revaux",
}
def correct_brand(raw):
    return BRAND_CORRECTIONS.get(str(raw).strip().lower(), str(raw).strip())

raw_brands = sorted(set(str(x).strip() for x in df["BrandD"].dropna().unique())) if "BrandD" in df.columns else []
matched = []
unmatched = []
for b in raw_brands:
    cleaned = clean_brand(correct_brand(b.title()))
    if cleaned in KNOWN_BRANDS:
        matched.append(f"{b} -> {cleaned}")
    else:
        unmatched.append(f"{b} -> {cleaned}")
print(f"Brands matched ({len(matched)}):")
for m in matched: print(f"  [OK] {m}")
print(f"Brands NOT in whitelist ({len(unmatched)}):")
for u in unmatched: print(f"  [MISS] {u}")
print()

# --- LenFilterCategory ANALYSIS (used to detect if we should switch to it) ---
if "LenFilterCategory" in df.columns:
    print(f"=== LenFilterCategory ANALYSIS ===")
    non_empty = df["LenFilterCategory"].dropna().astype(str).str.strip()
    non_empty = non_empty[non_empty.ne("")]
    print(f"Total rows: {len(df):,}")
    print(f"Rows with LenFilterCategory value: {len(non_empty):,}")
    print(f"Distinct values: {sorted(set(non_empty))[:20]}")

# --- Cross-check UV_TransparencyPerc range ---
if "UV_TransparencyPerc" in df.columns:
    print()
    print(f"=== UV_TransparencyPerc DISTRIBUTION ===")
    vals = pd.to_numeric(df["UV_TransparencyPerc"], errors="coerce").dropna()
    print(f"Total numeric: {len(vals):,}")
    print(f"  ==0:           {(vals == 0).sum():,}")
    print(f"  0<v<3:         {((vals > 0) & (vals < 3)).sum():,}")
    print(f"  3<=v<8 (Cat4): {((vals >= 3) & (vals < 8)).sum():,}")
    print(f"  8<=v<18 (Cat3):{((vals >= 8) & (vals < 18)).sum():,}")
    print(f"  18<=v<43 (Cat2):{((vals >= 18) & (vals < 43)).sum():,}")
    print(f"  43<=v<80 (Cat1):{((vals >= 43) & (vals < 80)).sum():,}")
    print(f"  80<=v<=100 (Cat0):{((vals >= 80) & (vals <= 100)).sum():,}")
