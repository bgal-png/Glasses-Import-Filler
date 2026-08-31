# Desktop app — where we got to, and what's left for you

Written Fri 2026-08-28. Everything below is committed and pushed to `main`.

## What's done (code complete, tested)

| Piece | State |
|---|---|
| `filler_core.py` — the fill engine, extracted UI-free | ✅ tested end-to-end without a DB |
| `admin_core.py` — every write operation, UI-free | ✅ 25 assertions passing |
| `desktop/` — PySide6 app, 6 tabs, dark mode, self-update | ✅ builds headless, all tabs wired |
| `data_source.py` — ETag + pickle snapshot cache | ✅ 21 assertions passing |
| `scripts/export_snapshot.py` + `publish-snapshot.yml` | ✅ written, **not yet run** |
| `app_manufacturer.py` — now a thin UI over the core | ⚠️ **please sanity-check once** (see below) |

Tabs: 🪄 Auto-Filler · 🔍 Barcode Checker · 🏭 Catalogue · ✏️ Rename · 📒 Registry.
The last three only appear when a `DB_URL` is set in Settings.

**Verified in real use (2026-08-31):** Auto-Filler end to end — opened a real
target file, filled it, saved, and the output kept its formatting — plus the
Barcode Checker. Catalogue / Rename / Registry are written and unit-tested but
have not been driven by hand yet.

**Removed:** the 🎨 Colours tab (fill missing colours from product photos). It
was never used successfully in the web version either — the photo filenames
never matched — and the underlying job turned out not to be needed, so the tab,
its shared logic and the equivalent Streamlit section were all deleted.

## ⚠️ One thing to verify first

I rewrote `app_manufacturer.py` to call the extracted engine instead of holding
the logic inline. The engine code was moved verbatim and I verified it fills
correctly against a real 14,077-row catalogue, but **the web filler itself
hasn't been run since**. Please do one normal fill on Streamlit Cloud and check
the output looks the same as before. If anything's off, `git revert 4c83ad1`
puts the old inline version back and nothing else breaks.

## What I could not do (needs your GitHub account)

### 1. Create the private data repo
New **private** repo, e.g. `bgal-png/Glasses-Filler-Data`. Empty is fine.
This is where the catalogue snapshot gets published. It must be private —
`master_catalog` is supplier master data and this code repo is public.

### 2. Two tokens (different scopes on purpose)

**a) Write token — for the Action that publishes the snapshot**
GitHub → Settings → Developer settings → Personal access tokens →
Fine-grained tokens → Generate new token
- Repository access: only `Glasses-Filler-Data`
- Permissions: **Contents: Read and write**

**b) Read-only token — for the desktop app**
Same place, a second token
- Repository access: only `Glasses-Filler-Data`
- Permissions: **Contents: Read-only**

This one goes into the .exe, so read-only matters: worst case a leak means
someone can read the product snapshot, not touch the database.

### 3. Three secrets on `Glasses-Import-Filler`
Settings → Secrets and variables → Actions:

| Secret | Value |
|---|---|
| `SNAPSHOT_REPO` | `bgal-png/Glasses-Filler-Data` |
| `SNAPSHOT_TOKEN` | the **write** token from 2a |
| `DB_URL` | already there |

### 4. Publish the first snapshot
Actions → **Publish catalogue snapshot** → Run workflow.
It then runs automatically after every Safilo/Luxottica auto-ingest and daily
at 06:45 UTC. Check the data repo afterwards — it should contain
`master_catalog.csv.gz`, `package_data.csv.gz`, `origin_data.csv.gz`,
`ingest_log.csv.gz` and `manifest.json`.

### 5. Point the app at it
Run the app, ⚙️ Settings:
- Snapshot repo: `bgal-png/Glasses-Filler-Data`
- Snapshot token: the **read-only** token from 2b
- Database URL: your `DB_URL` — only on your machine, unlocks the admin tabs
- Anthropic API key: optional, enables AI shape recognition

To bake the read-only snapshot settings into builds for colleagues, create
`desktop/defaults.py` (gitignored):

```python
SNAPSHOT_REPO = "bgal-png/Glasses-Filler-Data"
SNAPSHOT_TOKEN = "github_pat_…"   # the READ-ONLY one
```

Then colleagues get a working app with no setup, and no admin tabs.

## Running and building

```
"C:\gv\Scripts\python.exe" desktop\main.py
"C:\gv\Scripts\pyinstaller.exe" desktop\GlassesFiller.spec
```

Tests:
```
"C:\gv\Scripts\python.exe" desktop\main.py --selftest --force-admin
"C:\gv\Scripts\python.exe" desktop\test_data_source.py
python test_filler_core.py
python test_admin_core.py
```

## Still to do (phase 3)

1. **Drive the three admin tabs by hand** — Catalogue, Rename and Registry are
   unit-tested but nobody has clicked through them yet.
2. **Snapshot setup** (the section above) so startup stops re-downloading the
   whole catalogue from Supabase on every launch.
3. **Release flow** — bump `desktop/version.py`, tag `desktop-v1.0.0`, attach
   `GlassesFiller.exe` to a GitHub Release. Only then does self-update work.
