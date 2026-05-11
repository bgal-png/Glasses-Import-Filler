# -*- coding: utf-8 -*-
"""
Generic headless ingest — runs as a GitHub Action triggered by Apps Script
when a new file lands in a manufacturer's Drive `inbox` folder.

Supports any manufacturer in MANUFACTURER_CONFIG (currently safilo, luxottica,
marcolin, kering). Selects manufacturer from --mfg CLI arg or MFG_KEY env var.

Behavior:
  1. Reads service-account credentials from GDRIVE_SERVICE_ACCOUNT_JSON env var.
  2. Lists every file in the inbox folder (env: GDRIVE_INBOX_FOLDER_ID).
  3. For each file:
     - downloads it
     - if .zip: extracts and finds the .xlsx/.csv inside
     - runs `ingest.load_single_catalog` and `ingest.perform_upsert`
     - moves the ORIGINAL Drive file to the archive folder
  4. Logs everything to stdout (visible in GitHub Action logs).
  5. Exits non-zero if any file fails — but processes the rest of the batch first.

Environment variables expected:
    GDRIVE_SERVICE_ACCOUNT_JSON   Full JSON contents of the service account key.
    GDRIVE_INBOX_FOLDER_ID        Drive folder ID where files land.
    GDRIVE_ARCHIVE_FOLDER_ID      Drive folder ID where processed files are moved.
    DB_URL                        SQLAlchemy DSN for Supabase.
    MFG_KEY (optional)            Manufacturer key — overridable via --mfg arg.
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import tempfile
import traceback
import zipfile
from datetime import datetime, timezone

import pandas as pd
from sqlalchemy import create_engine

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# Local imports
_sys_path_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, _sys_path_root)
from dictionaries import MANUFACTURER_CONFIG  # noqa: E402
from ingest import load_single_catalog, perform_upsert  # noqa: E402


def _log(msg: str) -> None:
    """Stamped log line for GitHub Actions output."""
    print(f"[{datetime.now(timezone.utc).isoformat(timespec='seconds')}] {msg}", flush=True)


def _drive_client():
    raw_json = os.environ["GDRIVE_SERVICE_ACCOUNT_JSON"]
    info = json.loads(raw_json)
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/drive"],
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def _list_inbox_files(drive, inbox_folder_id: str) -> list[dict]:
    """Return all non-trashed files in the inbox folder."""
    files = []
    page_token = None
    while True:
        resp = drive.files().list(
            q=f"'{inbox_folder_id}' in parents and trashed = false",
            fields="nextPageToken, files(id, name, mimeType, size)",
            pageToken=page_token,
            pageSize=100,
        ).execute()
        files.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return files


def _download_file(drive, file_id: str, dest_path: str) -> None:
    request = drive.files().get_media(fileId=file_id)
    with open(dest_path, "wb") as f:
        downloader = MediaIoBaseDownload(f, request, chunksize=1024 * 1024)
        done = False
        while not done:
            _status, done = downloader.next_chunk()


def _move_to_archive(drive, file_id: str, inbox_folder_id: str, archive_folder_id: str) -> None:
    """Reparent a file from inbox to archive."""
    drive.files().update(
        fileId=file_id,
        addParents=archive_folder_id,
        removeParents=inbox_folder_id,
        fields="id, parents",
    ).execute()


def _extract_payload_from_zip(zip_path: str, dest_dir: str) -> str:
    """Open a ZIP, find the first .xlsx or .csv inside, extract it, return its path.
    Raises if there's no eligible payload."""
    with zipfile.ZipFile(zip_path) as zf:
        names = [n for n in zf.namelist() if not n.endswith("/")]
        eligible = [n for n in names if n.lower().endswith((".xlsx", ".csv"))]
        if not eligible:
            raise ValueError(
                f"ZIP contains no .xlsx/.csv files. Members: {names}"
            )
        # Pick the largest (in case of multiple — usually the real data file)
        eligible.sort(key=lambda n: zf.getinfo(n).file_size, reverse=True)
        target = eligible[0]
        zf.extract(target, dest_dir)
        return os.path.join(dest_dir, target)


def main(mfg: str) -> int:
    inbox_id = os.environ["GDRIVE_INBOX_FOLDER_ID"]
    archive_id = os.environ["GDRIVE_ARCHIVE_FOLDER_ID"]
    db_url = os.environ["DB_URL"]

    if mfg not in MANUFACTURER_CONFIG:
        _log(f"FATAL: unknown manufacturer key '{mfg}'. Known: {list(MANUFACTURER_CONFIG.keys())}")
        return 2

    _log(f"Starting ingest for mfg={mfg}")
    _log(f"Inbox folder:   {inbox_id}")
    _log(f"Archive folder: {archive_id}")

    drive = _drive_client()
    files = _list_inbox_files(drive, inbox_id)

    if not files:
        _log("No files in inbox. Nothing to do.")
        return 0

    _log(f"Found {len(files)} file(s) to process:")
    for f in files:
        _log(f"  - {f['name']} ({f.get('size', '?')} bytes, id={f['id']})")

    engine = create_engine(db_url, pool_pre_ping=True, pool_recycle=300)
    config = MANUFACTURER_CONFIG[mfg]

    failures = []
    successes = []

    for file_meta in files:
        file_id = file_meta["id"]
        name = file_meta["name"]
        _log("")
        _log(f"==== Processing: {name} ====")

        lower_name = name.lower()
        if not lower_name.endswith((".csv", ".xlsx", ".zip")):
            _log(f"  SKIPPED (unsupported file type)")
            continue

        with tempfile.TemporaryDirectory() as tmp:
            local_path = os.path.join(tmp, name)
            try:
                _log(f"  Downloading from Drive…")
                _download_file(drive, file_id, local_path)
                size_kb = os.path.getsize(local_path) / 1024
                _log(f"  Downloaded {size_kb:.1f} KB")

                # If ZIP, extract the data file inside
                if lower_name.endswith(".zip"):
                    _log(f"  Extracting ZIP…")
                    process_path = _extract_payload_from_zip(local_path, tmp)
                    _log(f"  Extracted: {os.path.basename(process_path)}")
                else:
                    process_path = local_path

                _log(f"  Running rules engine…")
                df, unmapped, skipped = load_single_catalog(mfg, config, process_path)
                _log(
                    f"  Engine output: {len(df):,} rows, "
                    f"{len(unmapped)} unmapped values, {len(skipped)} skipped 'NOT MAPPED'"
                )

                if df.empty:
                    _log(f"  WARN: engine returned empty DataFrame — moving file to archive anyway")
                else:
                    # Mirror app_admin's "expand by brands" behavior
                    expanded = pd.concat([df.copy() for _ in config["brands"]], ignore_index=True)
                    _log(f"  After brand expansion: {len(expanded):,} rows")
                    _log(f"  Upserting to master_catalog…")
                    msg = perform_upsert(expanded, engine)
                    _log(f"  Upsert: {msg}")

                _log(f"  Moving original Drive file to archive…")
                _move_to_archive(drive, file_id, inbox_id, archive_id)
                _log(f"  Done.")
                successes.append(name)

            except Exception as e:
                _log(f"  FAILED: {type(e).__name__}: {e}")
                _log(f"  Traceback:")
                for line in traceback.format_exc().splitlines():
                    _log(f"    {line}")
                failures.append((name, str(e)))

    _log("")
    _log(f"==== Summary ====")
    _log(f"Succeeded: {len(successes)}")
    for s in successes:
        _log(f"  + {s}")
    _log(f"Failed: {len(failures)}")
    for n, e in failures:
        _log(f"  - {n}: {e}")

    return 1 if failures else 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Headless manufacturer-catalog ingest.")
    parser.add_argument(
        "--mfg",
        default=os.environ.get("MFG_KEY", "safilo"),
        help="Manufacturer key (safilo, luxottica, kering, marcolin). "
             "Defaults to MFG_KEY env var or 'safilo'.",
    )
    args = parser.parse_args()
    try:
        sys.exit(main(args.mfg))
    except KeyError as e:
        _log(f"FATAL: required environment variable not set: {e}")
        sys.exit(2)
