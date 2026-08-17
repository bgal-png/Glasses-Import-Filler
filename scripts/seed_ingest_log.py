# -*- coding: utf-8 -*-
"""
One-time seed for the ingest_log table.

The 'last catalogue update' panel starts blank because ingest_log only began
recording when the feature was added. This backfills approximate last-update
dates (from git history) for producers set up before then, so the panel isn't
all 'never'. Only seeds producers NOT already present, so it never overwrites a
real timestamp written by a later ingest.

Run via the 'Seed ingest log' GitHub Action (uses the DB_URL secret), or
locally with DB_URL set in the environment.
"""
import os

import pandas as pd
from sqlalchemy import create_engine

# Approximate last-update timestamps (UTC ISO), derived from git history.
SEED = {
    "derigo": "2026-06-16T06:56:43+00:00",
    "thelios": "2026-07-17T11:39:12+00:00",
    "marcolin": "2026-07-17T05:29:04+00:00",
    "kering": "2026-02-24T10:34:00+00:00",
}


def main():
    db_url = os.environ["DB_URL"]
    engine = create_engine(db_url, pool_pre_ping=True)

    try:
        existing = pd.read_sql_table("ingest_log", con=engine)
    except Exception:
        existing = pd.DataFrame(columns=["manufacturer", "last_updated", "rows"])

    have = set(existing["manufacturer"].astype(str).str.lower()) if len(existing) else set()

    new_rows = []
    for mfg, ts in SEED.items():
        if mfg in have:
            print(f"skip {mfg} (already recorded — not overwriting)")
            continue
        new_rows.append({"manufacturer": mfg, "last_updated": ts, "rows": None})
        print(f"seed {mfg} -> {ts}")

    if new_rows:
        combined = pd.concat([existing, pd.DataFrame(new_rows)], ignore_index=True)
        combined.to_sql("ingest_log", con=engine, if_exists="replace", index=False)
        print(f"Done. Seeded {len(new_rows)} producer(s). Table now has {len(combined)} row(s).")
    else:
        print("Nothing to seed — all target producers already have timestamps.")


if __name__ == "__main__":
    main()
