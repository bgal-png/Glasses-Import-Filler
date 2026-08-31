# Desktop app — status

Last updated 2026-08-31. **v1.0.0 released.**

## Working and verified in real use

| Piece | State |
|---|---|
| 🪄 Auto-Filler | ✅ open → fill → save; output keeps its formatting |
| 🔍 Barcode Checker | ✅ file or pasted list |
| Snapshot data layer | ✅ live; second launch is instant from local cache |
| `filler_core.py` / `admin_core.py` | ✅ UI-free, unit-tested without a DB |
| `.exe` build | ✅ 85 MB, icon, released as `desktop-v1.0.0` |
| Self-update | ✅ updater reads the release; no false "update available" |

🏭 Catalogue · ✏️ Rename · 📒 Registry and the Streamlit filler were all
reported working by the user on 2026-08-31.

## Still outstanding

Nothing required for your own use — it's done and released.

**Only when you hand it to a colleague:**
1. Send them `GlassesFiller.exe` (or the Release link).
2. On their machine, open it once → ⚙️ Settings → paste the **snapshot repo** and
   the **read-only token** → Save. ~30 seconds, once.
3. They see only 🪄 Auto-Filler and 🔍 Barcode Checker, with no database
   credentials anywhere.

Deliberately **not** baking those into the build via `defaults.py`: a build with
the token in it must never be attached to a public Release, and distributing it
privately instead would break self-update. Token-free builds stay publishable
and auto-updating.

## Shipping an update

```
# 1. edit desktop/version.py   __version__ = "1.0.1"
"C:\gv\Scripts\pyinstaller.exe" desktop\GlassesFiller.spec
# 2. tag desktop-v1.0.1, new GitHub Release, attach dist\GlassesFiller.exe
```

Every installed copy checks on launch and offers it. The updater only looks at
tags starting `desktop-v`, and swaps the .exe via a .bat that waits for exit.

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
