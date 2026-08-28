# -*- coding: utf-8 -*-
"""Tests for the snapshot fetch / ETag / pickle-cache logic.

This is the part users would silently suffer from if it broke (stale or
unusable catalogue), and it can't be exercised by the UI selftest — so it gets
a real test with a stubbed HTTP layer. No network, no database, no Qt beyond
QSettings-free stub settings.

Run:  "C:\\gv\\Scripts\\python.exe" desktop\\test_data_source.py
"""
from __future__ import annotations

import gzip
import io
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_HERE)
for _p in (_HERE, _ROOT):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import pandas as pd  # noqa: E402

import data_source  # noqa: E402


class StubSettings:
    """Settings without QSettings/registry."""
    snapshot_repo = "acme/data"
    snapshot_token = "tok"
    snapshot_branch = "main"
    db_url = ""
    anthropic_key = ""


def _gz(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with gzip.open(buf, "wb") as fh:
        fh.write(df.to_csv(index=False).encode("utf-8"))
    return buf.getvalue()


MASTER = pd.DataFrame({
    "join_key": ["716736139197", "889214112798"],
    "Brand": ["Marc Jacobs", "Tom Ford"],
    "Assembled_Name": ["Marc Jacobs MARC 400 ISK", "Tom Ford FT0613 01D"],
    "Glasses_type": ["Frames", "Sunglasses"],
})
TINY = pd.DataFrame({"item_name": ["Carrera"], "country_master": ["Italy"]})
LOG = pd.DataFrame({"manufacturer": ["safilo"], "last_updated": ["2026-08-17T14:52:14+00:00"]})

BODIES = {
    "master_catalog.csv.gz": _gz(MASTER),
    "package_data.csv.gz": _gz(TINY),
    "origin_data.csv.gz": _gz(TINY),
    "ingest_log.csv.gz": _gz(LOG),
    "manifest.json": b'{"snapshot_format":1,"generated_utc":"2026-08-17T15:00:00+00:00"}',
}

calls = {"200": 0, "304": 0, "manifest": 0}


def make_fetch(mode: str):
    """mode: 'fresh' -> always 200; 'unchanged' -> 304 for tables; 'offline' -> raise"""
    def _fetch(url, token, etag=""):
        name = url.rsplit("/", 1)[-1]
        if mode == "offline":
            raise OSError("network is down")
        if name == "manifest.json":
            calls["manifest"] += 1
            return BODIES[name], '"m1"', 200
        if name not in BODIES:
            import urllib.error
            raise urllib.error.HTTPError(url, 404, "Not Found", None, None)
        if mode == "unchanged" and etag:
            calls["304"] += 1
            return None, etag, 304
        calls["200"] += 1
        return BODIES[name], f'"{name}-v1"', 200
    return _fetch


def check(label, cond, detail=""):
    print(f"  {'PASS' if cond else 'FAIL'}  {label}{(' — ' + detail) if detail and not cond else ''}")
    return bool(cond)


def main() -> int:
    ok = True
    data_source.clear_cache()
    settings = StubSettings()

    # ---------- 1. first load: downloads and parses ----------
    print("1. First load (cold cache, HTTP 200)")
    data_source._fetch = make_fetch("fresh")
    d1 = data_source.load_catalogue(settings)
    ok &= check("source is 'snapshot'", d1.source == "snapshot", d1.source)
    ok &= check("master has 2 rows", len(d1.master_db) == 2, str(len(d1.master_db)))
    ok &= check("origin loaded", len(d1.origin_df) == 1)
    ok &= check("ingest_log loaded", len(d1.ingest_log) == 1)
    ok &= check("generated_utc from manifest", d1.generated_utc.startswith("2026-08-17"), d1.generated_utc)
    ok &= check("barcodes stayed text", d1.master_db["join_key"].iloc[0] == "716736139197",
                repr(d1.master_db["join_key"].iloc[0]))
    ok &= check("4 files downloaded", calls["200"] == 4, str(calls["200"]))
    cache_files = sorted(os.listdir(data_source.cache_dir("snapshot")))
    ok &= check("gz + etag + pkl cached",
                all(any(f.endswith(ext) for f in cache_files) for ext in (".csv.gz", ".etag", ".pkl")),
                str(cache_files))

    # ---------- 2. second load: 304, served from pickle ----------
    print("2. Second load (server says unchanged -> pickle cache)")
    calls["200"] = calls["304"] = 0
    data_source._fetch = make_fetch("unchanged")
    d2 = data_source.load_catalogue(settings)
    ok &= check("nothing re-downloaded", calls["200"] == 0, f"200s={calls['200']}")
    ok &= check("all four got 304", calls["304"] == 4, str(calls["304"]))
    ok &= check("data identical", d2.master_db.equals(d1.master_db))
    ok &= check("still 'snapshot' (not stale)", d2.source == "snapshot", d2.source)

    # ---------- 3. offline: falls back to cache, flags it ----------
    print("3. Offline (network error -> cached copy, marked stale)")
    data_source._fetch = make_fetch("offline")
    d3 = data_source.load_catalogue(settings)
    ok &= check("still returns data", len(d3.master_db) == 2, str(len(d3.master_db)))
    ok &= check("source flagged 'cache'", d3.source == "cache", d3.source)
    ok &= check("told the user why", any("offline" in m for m in d3.messages), str(d3.messages))

    # ---------- 4. force refresh ignores the ETag ----------
    print("4. Force refresh (ignores ETag)")
    calls["200"] = calls["304"] = 0
    data_source._fetch = make_fetch("unchanged")
    data_source.load_catalogue(settings, force_refresh=True)
    ok &= check("re-downloaded despite 304 mode", calls["200"] == 4, str(calls["200"]))

    # ---------- 5. cold cache + offline = clear error ----------
    print("5. Cold cache + offline (must fail loudly, not silently)")
    data_source.clear_cache()
    data_source._fetch = make_fetch("offline")
    try:
        data_source.load_catalogue(settings)
        ok &= check("raises DataError", False, "no exception")
    except data_source.DataError as e:
        ok &= check("raises DataError", True)
        ok &= check("message mentions the table", "master_catalog" in str(e), str(e))

    # ---------- 6. nothing configured ----------
    print("6. No data source configured")
    data_source.clear_cache()
    blank = StubSettings()
    blank.snapshot_repo = ""
    try:
        data_source.load_catalogue(blank)
        ok &= check("raises DataError", False, "no exception")
    except data_source.DataError as e:
        ok &= check("raises DataError with guidance", "Settings" in str(e), str(e))

    data_source.clear_cache()
    print()
    print("ALL PASS" if ok else "FAILURES ABOVE")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
