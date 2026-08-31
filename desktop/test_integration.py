# -*- coding: utf-8 -*-
"""Integration test across the seam I hadn't covered:
data_source (snapshot) -> filler_core (fill).

Builds a real master_catalog by running the ingest engine over a catalogue file,
serves it through a stubbed HTTP layer exactly as the private snapshot repo
would, loads it via data_source, then fills a real target template. This proves
the DataFrame that data_source hands over is the shape fill_target expects.
"""
from __future__ import annotations

import gzip, io, os, sys

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_HERE)
for _p in (_HERE, _ROOT):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import pandas as pd
import data_source
from dictionaries import MANUFACTURER_CONFIG
from ingest import load_single_catalog
from filler_core import FillOptions, changed_columns, fill_target, read_target_file

CATALOGUE = r"C:\Users\blank\Downloads\SafiloAvailability_CZ0120 (11).csv"
TEMPLATE = r"C:\Users\blank\Desktop\OLD PC\BG Old\Glasses Imports\filler-testing.xlsx"

ok = True
def check(label, cond, detail=""):
    global ok
    print(f"  {'PASS' if cond else 'FAIL'}  {label}" + (f" -- {detail}" if detail and not cond else ""))
    ok = ok and bool(cond)

print("Building a real master_catalog via the ingest engine...")
mdf, _u, _s = load_single_catalog("safilo", MANUFACTURER_CONFIG["safilo"], CATALOGUE)
print(f"  {len(mdf):,} rows")

def gz(df):
    buf = io.BytesIO()
    with gzip.open(buf, "wb") as fh:
        fh.write(df.to_csv(index=False).encode("utf-8"))
    return buf.getvalue()

BODIES = {
    "master_catalog.csv.gz": gz(mdf),
    "package_data.csv.gz": gz(pd.DataFrame({"item_name": ["Carrera"], "case_length": ["160"],
                                            "case_height": ["60"], "case_width": ["70"],
                                            "case_weight": ["90"], "item_weight": ["30"]})),
    "origin_data.csv.gz": gz(pd.DataFrame({"item_name": ["Carrera"], "country_master": ["Italy"]})),
    "ingest_log.csv.gz": gz(pd.DataFrame({"manufacturer": ["safilo"],
                                          "last_updated": ["2026-08-28T10:00:00+00:00"]})),
    "manifest.json": b'{"snapshot_format":1,"generated_utc":"2026-08-28T10:00:00+00:00"}',
}

def fake_fetch(url, token, etag=""):
    name = url.rsplit("/", 1)[-1]
    if name not in BODIES:
        import urllib.error
        raise urllib.error.HTTPError(url, 404, "Not Found", None, None)
    return BODIES[name], f'"{name}-v1"', 200

class S:
    snapshot_repo, snapshot_token, snapshot_branch = "acme/data", "tok", "main"
    db_url = anthropic_key = ""

print("Loading it back through data_source (as the .exe would)...")
data_source.clear_cache()
data_source._fetch = fake_fetch
cat = data_source.load_catalogue(S())
check("row count survives the round trip", len(cat.master_db) == len(mdf), f"{len(cat.master_db)} vs {len(mdf)}")
check("package/origin/log loaded", not cat.package_df.empty and not cat.origin_df.empty and not cat.ingest_log.empty)

print("Filling a real target template with it...")
template = read_target_file(TEMPLATE)
keys = list(mdf[mdf["Glasses_type"] == "Sunglasses"]["join_key"].head(5)) + \
       list(mdf[mdf["Glasses_type"] == "Frames"]["join_key"].head(5))
raw = list(mdf.set_index("join_key").loc[keys, "Barcode"])
target = pd.DataFrame(columns=template.columns)
target["Barcode"] = raw
target = target.astype(object)
original = target.copy()

filled, report = fill_target(target, cat.master_db, cat.package_df, cat.origin_df,
                             options=FillOptions(priv_sun="1001", priv_eye="2001"))
check("every barcode matched", report.match_count == len(target), f"{report.match_count}/{len(target)}")
cols = changed_columns(original, filled)
check("many columns filled", len(cols) > 25, len(cols))
check("case dims came from package_data",
      any(str(v).strip() for v in filled["Case length (mm)"]), "all empty")
check("origin came from origin_data",
      any(str(v).strip() for v in filled.get("Item origin country", pd.Series(dtype=object))), "all empty")
check("no float-formatted sizes (e.g. '54.0')",
      not any(str(v).endswith(".0") for v in filled["Combination (size on glasses)"] if str(v).strip()),
      [v for v in filled["Combination (size on glasses)"] if str(v).endswith(".0")][:3])

data_source.clear_cache()
print()
print("ALL PASS" if ok else "FAILURES ABOVE")
sys.exit(0 if ok else 1)
