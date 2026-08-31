# -*- coding: utf-8 -*-
"""Investigate which column truly contains the sunglass filter category."""
import pandas as pd
import glob

FILES = sorted(glob.glob(r"C:\Users\blank\Downloads\SafiloAvailability_CZ0120*.csv"))
dfs = []
for f in FILES:
    d = pd.read_csv(f, dtype=str, on_bad_lines="skip", sep=";")
    if len(d.columns) <= 1:
        d = pd.read_csv(f, dtype=str, on_bad_lines="skip", sep=",")
    d.columns = d.columns.astype(str).str.strip()
    dfs.append(d)
df = pd.concat(dfs, ignore_index=True)
print(f"Total rows: {len(df):,}")
print()

# Focus on SUNGLASS frames only
sun = df[df["TypeD"].astype(str).str.strip().eq("Sunglass frames")].copy()
print(f"Sunglass frames: {len(sun):,}")
print()

print("=" * 60)
print("LenFilterCategory distribution for SUNGLASS frames")
print("=" * 60)
counts = sun["LenFilterCategory"].fillna("(empty)").value_counts(dropna=False)
print(counts.head(30).to_string())
print()

print("=" * 60)
print("UV_TransparencyPerc distribution for SUNGLASS frames")
print("=" * 60)
sun_uvt = pd.to_numeric(sun["UV_TransparencyPerc"], errors="coerce")
print(f"  ==0:               {(sun_uvt == 0).sum():,}")
print(f"  0<v<3:             {((sun_uvt > 0) & (sun_uvt < 3)).sum():,}")
print(f"  3<=v<8 (Cat4 VLT): {((sun_uvt >= 3) & (sun_uvt < 8)).sum():,}")
print(f"  8<=v<18 (Cat3):    {((sun_uvt >= 8) & (sun_uvt < 18)).sum():,}")
print(f"  18<=v<43 (Cat2):   {((sun_uvt >= 18) & (sun_uvt < 43)).sum():,}")
print(f"  43<=v<80 (Cat1):   {((sun_uvt >= 43) & (sun_uvt < 80)).sum():,}")
print(f"  80<=v<=100 (Cat0): {((sun_uvt >= 80) & (sun_uvt <= 100)).sum():,}")
print()

print("=" * 60)
print("Cross-tab: LenFilterCategory vs UV_TransparencyPerc range")
print("=" * 60)
def bucket(v):
    try:
        v = float(v)
    except:
        return "NaN"
    if v == 0: return "0"
    if v < 3: return "0-3"
    if v < 8: return "3-8"
    if v < 18: return "8-18"
    if v < 43: return "18-43"
    if v < 80: return "43-80"
    return "80+"
sun_copy = sun.copy()
sun_copy["uvt_bucket"] = sun_copy["UV_TransparencyPerc"].apply(bucket)
sun_copy["lfc"] = sun_copy["LenFilterCategory"].fillna("(empty)").astype(str)
ct = pd.crosstab(sun_copy["lfc"], sun_copy["uvt_bucket"])
print(ct.head(20).to_string())
print()

# Look at LenFilterCategory for OPTICAL frames (should be empty/N/A)
print("=" * 60)
print("LenFilterCategory for OPTICAL frames (sanity check — should be empty)")
print("=" * 60)
opt = df[df["TypeD"].astype(str).str.strip().eq("Optical frames")]
print(opt["LenFilterCategory"].fillna("(empty)").value_counts(dropna=False).head(10).to_string())
