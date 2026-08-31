import pandas as pd
from dictionaries import MANUFACTURER_CONFIG
from ingest import load_single_catalog

FILE = r"C:\Users\blank\Downloads\SafiloAvailability_CZ0120 (1).csv"
cfg = MANUFACTURER_CONFIG["safilo"]
df, _u, _s = load_single_catalog("safilo", cfg, FILE)

print("Sample Glasses_color_code values (first 15 distinct):")
samples = df["Glasses_color_code"].dropna().unique()[:15]
for s in samples:
    print(f"  '{s}'")
print()
print(f"Total distinct: {df['Glasses_color_code'].nunique()}")
print(f"Rows with code: {(df['Glasses_color_code'] != '').sum():,} / {len(df):,}")
print(f"Rows with '/' (both frame+lens): {df['Glasses_color_code'].str.contains('/', na=False).sum():,}")
print(f"Rows without '/' (frame only or lens only): {((df['Glasses_color_code'] != '') & ~df['Glasses_color_code'].str.contains('/', na=False)).sum():,}")
