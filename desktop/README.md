# Glasses Filler — desktop app

PySide6 desktop version of the filler tools, following the same house pattern as
[Glasses-Validator-Desktop](https://github.com/bgal-png/Glasses-Validator-Desktop).

## Why it lives in this repo (not its own)

The validator has a separate desktop repo, but this tool's domain logic
(`ingest.py`, `dictionaries.py`, `filler_core.py` — ~4,000 lines) changes every
time a manufacturer changes their file format or a colour keyword is added.
Keeping the desktop app in the same repo means one source of truth instead of a
permanent copy-and-resync chore. Streamlit Cloud ignores this folder.

## Architecture

```
repo root                 shared, UI-free logic (also used by the web apps)
  dictionaries.py         mappings, colour classifier, brand tables
  ingest.py               per-manufacturer catalogue parsers
  filler_core.py          the auto-fill engine  ← extracted in 4c83ad1

desktop/
  main.py                 QMainWindow: tabs + toolbar + right dock + dark mode
  version.py              __version__, release repo/tag prefix
  settings.py             QSettings wrapper (credentials policy lives here)
  settings_dialog.py      ⚙️ Settings
  data_source.py          snapshot fetch (ETag + pickle cache) / direct-DB fallback
  workers.py              QThread worker (done/failed/progress)
  updater.py              self-update from GitHub Releases
  theme.py                dark palette + the shared house colours
  widgets.py              MetricCard/MetricRow, DataFrameTable (copyable headers)
  tabs/                   one QWidget per tab, common BaseTab contract
```

Nothing slow runs on the UI thread; every long job goes through `Worker`.

## Where the data comes from

The .exe must not carry database credentials, and pulling 50k+ rows from
Supabase per user per session would blow the egress budget. So:

1. **Snapshot (normal path).** The `publish-snapshot` Action exports
   `master_catalog`, `package_data`, `origin_data` and `ingest_log` to gzipped
   CSV in a **private** data repo after every ingest. The app fetches them with
   an HTTP **ETag** conditional request and keeps a **pickle** of the parsed
   frame in `%LOCALAPPDATA%\GlassesFiller\snapshot`. Unchanged data = nothing
   downloaded and nothing parsed. Offline = last cached copy, clearly flagged.
2. **Direct database (admin fallback).** If no snapshot is configured but a
   `DB_URL` is set in Settings, the tables are read straight from Supabase.

## Credentials policy

| Setting | Shippable in a build? | Grants |
|---|---|---|
| `snapshot_repo` / `snapshot_token` | Yes (see `defaults.py`) | read the catalogue snapshot |
| `db_url` | **Never** | read/write the live database |
| `anthropic_key` | No | AI shape recognition, billed to that key |

Admin tabs stay hidden until a `DB_URL` is present, so **one build serves
everyone**: colleagues get Auto-Filler + Barcode Checker, whoever pastes a
`DB_URL` gets the admin tabs too.

To ship read-only defaults, create an uncommitted `desktop/defaults.py`:

```python
SNAPSHOT_REPO = "bgal-png/Glasses-Filler-Data"
SNAPSHOT_TOKEN = "github_pat_…"   # read-only, that repo only
```

It is in `.gitignore` — never commit a real token.

## Run from source

The Microsoft Store Python's path is long enough that installing PySide6 fails
with `OSError [Errno 2] … enable-long-paths`, so use the short-path venv:

```
"C:\gv\Scripts\python.exe" desktop\main.py
```

## Build the .exe

```
"C:\gv\Scripts\pyinstaller.exe" desktop\GlassesFiller.spec
```

Produces `dist\GlassesFiller.exe`. Needs nothing installed on the target
machine; it is large (~200 MB) and unsigned, so SmartScreen warns once.

## Release (enables self-update)

1. Bump `__version__` in `desktop/version.py`.
2. Tag `desktop-v<x.y.z>` and create a GitHub Release on this repo.
3. Attach `GlassesFiller.exe` to it.

Installed copies compare their version with the newest `desktop-v*` tag on
launch and offer to update.

## Tests

```
"C:\gv\Scripts\python.exe" desktop\main.py --selftest   # builds the whole UI headless
"C:\gv\Scripts\python.exe" desktop\test_data_source.py  # ETag / cache / offline logic
python test_filler_core.py                              # fill engine, no DB needed
```
