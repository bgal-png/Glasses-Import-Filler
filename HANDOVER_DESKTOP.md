# Desktop app — status

Last updated 2026-08-31.

## Working and verified in real use

| Piece | State |
|---|---|
| 🪄 Auto-Filler | ✅ open → fill → save; output keeps its formatting |
| 🔍 Barcode Checker | ✅ file or pasted list |
| Snapshot data layer | ✅ live; second launch is instant from local cache |
| `filler_core.py` / `admin_core.py` | ✅ UI-free, unit-tested without a DB |
| `.exe` build | ✅ 85 MB, runs, all tabs build (throwaway build — not released) |

## Written but never clicked through

| Tab | Notes |
|---|---|
| 🏭 Catalogue | ingest + typed-confirmation delete. Writes to the live catalogue — try it on one manufacturer first. |
| ✏️ Rename | barcode → name list |
| 📒 Registry | store filled files, check barcodes against them |

## Still outstanding

1. **One normal fill on the Streamlit filler.** `app_manufacturer.py` was rewritten
   to call the extracted engine (commit `4c83ad1`). The engine is tested three
   ways, but the web app hasn't been run since. `git revert 4c83ad1` restores the
   old inline version if anything looks wrong.
2. **Drive the three admin tabs by hand.**
3. **Build + Release**, when you want a distributable .exe:
   - bump `__version__` in `desktop/version.py`
   - `"C:\gv\Scripts\pyinstaller.exe" desktop\GlassesFiller.spec`
   - tag `desktop-v<x.y.z>`, create a GitHub Release, attach `GlassesFiller.exe`
   - only then does in-app self-update work
4. **`desktop/defaults.py`** before handing the .exe to colleagues, so they need no
   setup. Gitignored — never commit a real token:
   ```python
   SNAPSHOT_REPO = "bgal-png/Glasses-Filler-Data"
   SNAPSHOT_TOKEN = "github_pat_…"   # the READ-ONLY token
   ```
   Colleagues then get Auto-Filler + Barcode Checker only, with no database
   credentials anywhere.

## How the data flows

```
catalogue file ─► ingest (rules engine) ─► Supabase master_catalog
                                                  │
                        publish-snapshot Action ───┘   (after every auto-ingest,
                                 │                      daily 06:45 UTC, or manual)
                                 ▼
                  private repo  bgal-png/Glasses-Filler-Data
                        master_catalog.csv.gz + 3 more + manifest.json
                                 │  ETag conditional request
                                 ▼
                  desktop app ── %LOCALAPPDATA%\GlassesFiller  (parsed pickle cache)
```

Unchanged data downloads nothing and parses nothing. Offline falls back to the
cached copy and says so in the status bar.

**Consequence to remember:** the filler reads the *snapshot*, not the database.
After a **manual** admin write the app now reminds you to re-run
*Publish catalogue snapshot*; automatic ingests trigger it themselves.

## Credentials

| Setting | Where it lives | Grants |
|---|---|---|
| `SNAPSHOT_TOKEN` (write) | GitHub Actions secret | push snapshots to the data repo |
| snapshot read-only token | app Settings / `defaults.py` | read the snapshot |
| `DB_URL` | app Settings, your machine only | read/write the live database; unlocks the admin tabs |
| Anthropic key | app Settings, optional | AI shape recognition |

## Commands

```
"C:\gv\Scripts\python.exe" desktop\main.py
"C:\gv\Scripts\pyinstaller.exe" desktop\GlassesFiller.spec

"C:\gv\Scripts\python.exe" desktop\main.py --selftest --force-admin
"C:\gv\Scripts\python.exe" desktop\test_data_source.py
"C:\gv\Scripts\python.exe" desktop\test_integration.py
python test_filler_core.py
python test_admin_core.py
```
