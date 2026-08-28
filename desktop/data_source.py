# -*- coding: utf-8 -*-
"""Catalogue data loading for the desktop app.

Two sources, in priority order:

1. **Snapshot** (normal path, no DB credentials): gzipped CSVs published to a
   private data repo by the `publish-snapshot` GitHub Action. Fetched with an
   HTTP **ETag** conditional request; the *parsed* frame is cached as a pickle
   in %LOCALAPPDATA%\\GlassesFiller\\snapshot. So a launch with unchanged data
   downloads nothing and parses nothing.
2. **Direct database** (admin fallback): if no snapshot is configured but a
   DB_URL is present in settings, read the tables straight from Supabase. This
   makes the app usable before the snapshot repo exists, and is the same code
   path the web app uses.

If the network fails, a previously cached snapshot is used and the caller is
told the data may be stale — the app must never become unusable offline.
"""
from __future__ import annotations

import gzip
import io
import json
import os
import pickle
import urllib.error
import urllib.request
from dataclasses import dataclass, field
from typing import Callable, Optional

import pandas as pd

from app_paths import cache_dir

TABLES = ["master_catalog", "package_data", "origin_data", "ingest_log"]
_TIMEOUT = 60


class DataError(Exception):
    """Raised when no data could be obtained from any source."""


@dataclass
class CatalogueData:
    master_db: pd.DataFrame = field(default_factory=pd.DataFrame)
    package_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    origin_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    ingest_log: pd.DataFrame = field(default_factory=pd.DataFrame)
    source: str = "unknown"          # snapshot | cache | database
    generated_utc: str = ""
    messages: list = field(default_factory=list)

    @property
    def is_empty(self) -> bool:
        return self.master_db is None or self.master_db.empty


ProgressFn = Optional[Callable[[float, str], None]]


def _tick(progress: ProgressFn, frac: float, text: str) -> None:
    if progress is not None:
        try:
            progress(max(0.0, min(1.0, frac)), text)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Snapshot fetching
# ---------------------------------------------------------------------------

def _snap_paths(table: str) -> tuple:
    d = cache_dir("snapshot")
    return (
        os.path.join(d, f"{table}.csv.gz"),
        os.path.join(d, f"{table}.etag"),
        os.path.join(d, f"{table}.pkl"),
    )


def _read_text(path: str) -> str:
    try:
        with open(path, "r", encoding="utf-8") as fh:
            return fh.read().strip()
    except Exception:
        return ""


def _raw_url(repo: str, branch: str, filename: str) -> str:
    return f"https://raw.githubusercontent.com/{repo}/{branch}/{filename}"


def _fetch(url: str, token: str, etag: str = "") -> tuple:
    """GET url with optional token + If-None-Match.
    Returns (body_bytes_or_None, new_etag, status). None body means 304."""
    req = urllib.request.Request(url)
    req.add_header("User-Agent", "GlassesFiller-Desktop")
    if token:
        req.add_header("Authorization", f"Bearer {token}")
    if etag:
        req.add_header("If-None-Match", etag)
    try:
        with urllib.request.urlopen(req, timeout=_TIMEOUT) as resp:
            return resp.read(), resp.headers.get("ETag", ""), resp.status
    except urllib.error.HTTPError as e:
        if e.code == 304:  # urllib treats "not modified" as an error
            return None, etag, 304
        raise


def _parse_csv_gz(body: bytes) -> pd.DataFrame:
    """Everything as text — matches how the filler treats every value, and
    keeps barcodes/sizes from being coerced to numbers."""
    with gzip.open(io.BytesIO(body), "rb") as fh:
        return pd.read_csv(fh, dtype=str, keep_default_na=False, na_values=[""])


def _load_table_from_snapshot(
    table: str, repo: str, token: str, branch: str, force: bool
) -> tuple:
    """Return (DataFrame, note). Uses ETag + pickle cache; falls back to the
    cached copy if the network is unavailable."""
    gz_path, etag_path, pkl_path = _snap_paths(table)
    etag = "" if force else _read_text(etag_path)
    url = _raw_url(repo, branch, f"{table}.csv.gz")

    try:
        body, new_etag, status = _fetch(url, token, etag)
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return pd.DataFrame(), f"{table}: not in snapshot (404)"
        if os.path.exists(pkl_path):
            return _unpickle(pkl_path), f"{table}: HTTP {e.code} — using cached copy"
        raise DataError(f"{table}: HTTP {e.code} fetching snapshot and no cache available") from e
    except Exception as e:
        if os.path.exists(pkl_path):
            return _unpickle(pkl_path), f"{table}: offline — using cached copy"
        raise DataError(f"{table}: {e} and no cache available") from e

    if status == 304 and os.path.exists(pkl_path):
        return _unpickle(pkl_path), ""

    if body is None:  # 304 but the parsed cache is gone — reparse stored gz
        if os.path.exists(gz_path):
            with open(gz_path, "rb") as fh:
                body = fh.read()
        else:
            raise DataError(f"{table}: server says unchanged but no local copy exists")
    else:
        with open(gz_path, "wb") as fh:
            fh.write(body)
        if new_etag:
            with open(etag_path, "w", encoding="utf-8") as fh:
                fh.write(new_etag)

    df = _parse_csv_gz(body)
    try:
        with open(pkl_path, "wb") as fh:
            pickle.dump(df, fh, protocol=pickle.HIGHEST_PROTOCOL)
    except Exception:
        pass  # a broken cache write must not break loading
    return df, ""


def _unpickle(path: str) -> pd.DataFrame:
    try:
        with open(path, "rb") as fh:
            return pickle.load(fh)
    except Exception:
        return pd.DataFrame()


def _load_manifest(repo: str, token: str, branch: str) -> dict:
    try:
        body, _etag, _status = _fetch(_raw_url(repo, branch, "manifest.json"), token)
        if body:
            data = json.loads(body.decode("utf-8"))
            with open(os.path.join(cache_dir("snapshot"), "manifest.json"), "wb") as fh:
                fh.write(body)
            return data
    except Exception:
        pass
    try:
        with open(os.path.join(cache_dir("snapshot"), "manifest.json"), "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def load_catalogue(settings, progress: ProgressFn = None, force_refresh: bool = False) -> CatalogueData:
    repo = settings.snapshot_repo
    token = settings.snapshot_token
    branch = settings.snapshot_branch
    db_url = settings.db_url

    if repo:
        _tick(progress, 0.05, "Checking catalogue snapshot…")
        out = CatalogueData(source="snapshot")
        manifest = _load_manifest(repo, token, branch)
        out.generated_utc = str(manifest.get("generated_utc", ""))

        frames = {}
        for i, table in enumerate(TABLES, start=1):
            _tick(progress, 0.05 + 0.85 * (i - 1) / len(TABLES), f"Loading {table}…")
            try:
                df, note = _load_table_from_snapshot(table, repo, token, branch, force_refresh)
            except DataError as e:
                if table == "master_catalog":
                    raise
                df, note = pd.DataFrame(), f"{table}: {e}"
            frames[table] = df
            if note:
                out.messages.append(note)
                if "cached copy" in note or "offline" in note:
                    out.source = "cache"

        out.master_db = frames.get("master_catalog", pd.DataFrame())
        out.package_df = frames.get("package_data", pd.DataFrame())
        out.origin_df = frames.get("origin_data", pd.DataFrame())
        out.ingest_log = frames.get("ingest_log", pd.DataFrame())

        if out.is_empty:
            raise DataError(
                "The catalogue snapshot is empty. Check that the publish-snapshot "
                "Action has run and that the snapshot repo/token are correct."
            )
        _tick(progress, 1.0, f"Catalogue ready ({len(out.master_db):,} products).")
        return out

    if db_url:
        return _load_from_database(db_url, progress)

    raise DataError(
        "No data source configured.\n\n"
        "Set the catalogue snapshot repo + token in Settings (normal use), or "
        "paste a DB_URL to read the database directly (admin)."
    )


def _load_from_database(db_url: str, progress: ProgressFn = None) -> CatalogueData:
    from sqlalchemy import create_engine

    _tick(progress, 0.05, "Connecting to the database…")
    out = CatalogueData(source="database")
    engine = create_engine(db_url, pool_pre_ping=True, pool_recycle=300)

    _tick(progress, 0.2, "Loading master catalogue…")
    try:
        out.master_db = pd.read_sql_table("master_catalog", con=engine)
    except Exception as e:
        raise DataError(f"Could not read master_catalog: {e}") from e

    for frac, table, attr in (
        (0.7, "package_data", "package_df"),
        (0.8, "origin_data", "origin_df"),
        (0.9, "ingest_log", "ingest_log"),
    ):
        _tick(progress, frac, f"Loading {table}…")
        try:
            setattr(out, attr, pd.read_sql_table(table, con=engine))
        except Exception:
            pass

    if out.is_empty:
        raise DataError("master_catalog is empty — upload a catalogue first.")
    _tick(progress, 1.0, f"Catalogue ready ({len(out.master_db):,} products).")
    return out


def cache_size_bytes() -> int:
    total = 0
    d = cache_dir("snapshot")
    for name in os.listdir(d):
        try:
            total += os.path.getsize(os.path.join(d, name))
        except Exception:
            pass
    return total


def clear_cache() -> None:
    d = cache_dir("snapshot")
    for name in os.listdir(d):
        try:
            os.remove(os.path.join(d, name))
        except Exception:
            pass
