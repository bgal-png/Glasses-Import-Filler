# -*- coding: utf-8 -*-
"""
Export the catalogue tables to a compact snapshot for the desktop app.

The desktop app must NOT hold database credentials, and pulling 50k+ rows from
Supabase per user per session would blow the egress budget. So a GitHub Action
runs this after every ingest and publishes the tables to a PRIVATE data repo;
the desktop app fetches them with an ETag conditional request and caches a
parsed copy locally (the same pattern as the Glasses Validator Desktop).

Wire format is gzipped CSV, deliberately:
  * no pyarrow dependency, so the packaged .exe stays ~40 MB smaller and
    avoids PyInstaller hidden-import problems;
  * everything is read back with dtype=str, which is what the filler wants
    anyway (no int/float coercion on barcodes or sizes);
  * the desktop keeps a local *pickle* of the parsed frame, so the CSV is only
    ever parsed when the data actually changed — parsing, not downloading, is
    the slow part.
Never pickle across the wire: pandas versions differ between the Action runner
and the packaged .exe.

Usage:
    DB_URL=... python scripts/export_snapshot.py --out ./snapshot
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime, timezone

import pandas as pd
from sqlalchemy import create_engine

# Tables the desktop app needs. master_catalog is the big one; the rest are tiny.
TABLES = ["master_catalog", "package_data", "origin_data", "ingest_log"]

SNAPSHOT_FORMAT = 1  # bump if the desktop app must reject older snapshots


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="snapshot", help="output directory")
    args = ap.parse_args()

    db_url = os.environ.get("DB_URL")
    if not db_url:
        print("FATAL: DB_URL environment variable is not set.", file=sys.stderr)
        return 2

    os.makedirs(args.out, exist_ok=True)
    engine = create_engine(db_url, pool_pre_ping=True)

    manifest = {
        "snapshot_format": SNAPSHOT_FORMAT,
        "generated_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "tables": {},
    }

    for table in TABLES:
        try:
            df = pd.read_sql_table(table, con=engine)
        except Exception as e:
            print(f"  {table}: SKIPPED ({type(e).__name__}: {e})")
            continue

        path = os.path.join(args.out, f"{table}.csv.gz")
        df.to_csv(path, index=False, compression="gzip", encoding="utf-8")
        size_kb = os.path.getsize(path) / 1024
        manifest["tables"][table] = {"rows": int(len(df)), "columns": int(len(df.columns))}
        print(f"  {table}: {len(df):,} rows x {len(df.columns)} cols -> {size_kb:,.0f} KB")

    if not manifest["tables"]:
        print("FATAL: no tables exported.", file=sys.stderr)
        return 1

    with open(os.path.join(args.out, "manifest.json"), "w", encoding="utf-8") as fh:
        json.dump(manifest, fh, indent=2)

    total_kb = sum(
        os.path.getsize(os.path.join(args.out, f)) for f in os.listdir(args.out)
    ) / 1024
    print(f"Snapshot written to {args.out} ({total_kb:,.0f} KB total)")
    print(f"generated_utc = {manifest['generated_utc']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
