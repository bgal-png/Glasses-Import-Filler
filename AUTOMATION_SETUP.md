# Safilo Drive → Auto-Ingest Setup

End-to-end automation: drop a Safilo CSV in a Drive folder, it's ingested into Supabase within ~5 minutes.

## Architecture

```
Drive `inbox/` folder
   │  (new file lands)
   ▼
Apps Script (polls every 5 min, fires on first non-empty poll)
   │  HTTP POST to GitHub API: repository_dispatch
   ▼
GitHub Actions: .github/workflows/safilo-ingest.yml
   │  pip install, run scripts/auto_ingest_safilo.py
   ▼
Python script
   │  ① download every file from Drive `inbox/`
   │  ② process via ingest.load_single_catalog (same code as admin app)
   │  ③ upsert into Supabase `master_catalog`
   │  ④ move file from `inbox/` to `archive/`
   ▼
Done — file now in archive/, DB updated.
```

A safety-net daily cron at 06:00 UTC also runs the script.
Manual trigger always available from GitHub Actions tab.

---

## Step 2 — GCP project + service account

1. Open [console.cloud.google.com](https://console.cloud.google.com)
2. Top bar → project dropdown → **New Project**. Name it e.g. `glasses-import-automation`. Create.
3. Make sure that new project is selected (top bar shows its name).
4. Left sidebar → **APIs & Services → Library** → search "Google Drive API" → **Enable**.
5. Left sidebar → **APIs & Services → Credentials** → **Create Credentials** → **Service account**.
   - Name: `safilo-ingest-bot` (or anything)
   - Service account ID auto-fills
   - Skip the optional grant-access steps → **Done**
6. Click the new service account row → **Keys** tab → **Add Key** → **Create new key** → **JSON** → **Create**. A `.json` file downloads. **Keep it safe**.
7. Copy the service account's **email address** (looks like `safilo-ingest-bot@your-project.iam.gserviceaccount.com`). You'll need it next.

## Step 3 — Drive folders

1. In Google Drive, create a folder named `Safilo Import` (or whatever).
2. Inside it, create two subfolders: `inbox` and `archive`.
3. Right-click `inbox` → **Share** → paste the service account email → set role to **Editor** → **Send** (uncheck "notify"). Repeat for `archive`.
4. Open `inbox` → look at the URL: `drive.google.com/drive/folders/XXXXXXXXX` — copy that `XXXXXXXXX`. That's your `INBOX_FOLDER_ID`.
5. Same for `archive` → `ARCHIVE_FOLDER_ID`.

## Step 4 — GitHub secrets

In the repo on github.com → **Settings → Secrets and variables → Actions → New repository secret**. Add four:

| Secret name | Value |
|---|---|
| `GDRIVE_SERVICE_ACCOUNT_JSON` | Open the `.json` file from Step 2.6 in any text editor and paste the **entire contents** (it's a JSON object) |
| `GDRIVE_INBOX_FOLDER_ID` | From Step 3.4 |
| `GDRIVE_ARCHIVE_FOLDER_ID` | From Step 3.5 |
| `DB_URL` | Same Supabase DSN as `st.secrets["DB_URL"]` |

## Step 5 — GitHub Personal Access Token (for Apps Script)

1. github.com → top right avatar → **Settings → Developer settings → Personal access tokens → Tokens (classic) → Generate new token (classic)**.
2. Name: `apps-script-safilo-dispatch`. Expiry: 1 year (or no expiry).
3. Scope: check **`repo`** (full control of private repos). That's the only one needed.
4. **Generate token** → copy the token string (`ghp_...`). You only see it once.

## Step 6 — Apps Script

1. Open [script.google.com](https://script.google.com) → **New project**.
2. Delete the empty `Code.gs` content. Paste the entire contents of `apps_script/drive_watcher.gs` from this repo.
3. At the top of the file, fill in:
   - `INBOX_FOLDER_ID` — from Step 3.4
   - `GITHUB_OWNER` — already filled (`bgal-png`)
   - `GITHUB_REPO` — already filled (`Glasses-Import-Filler`)
4. Left sidebar → **Project Settings (⚙ icon) → Script Properties → Add script property**:
   - Property name: `GITHUB_PAT`
   - Value: the token from Step 5.4
   - **Save**
5. Back in the code editor, the top bar dropdown should show `pollInboxAndDispatch`. Change it to `setupTrigger` and click **Run**.
6. Authorize when prompted (Google warns about an "unverified app" — that's your own script; click **Advanced → Go to (project name)**).
7. Check Apps Script left sidebar → **Triggers (clock icon)** → you should see one trigger pointing at `pollInboxAndDispatch`, "every 5 minutes".

## Step 7 — Test end-to-end

1. Drop one of the daily Safilo CSVs into the `inbox` folder in Drive.
2. Wait up to 5 minutes (or in Apps Script → run `testDispatchNow` to fire immediately).
3. Open the repo on github.com → **Actions** tab. You should see a new "Safilo Drive ingest" run in progress.
4. Click into it → expand "Run ingest" to see live logs.
5. When it finishes successfully:
   - The file should now be in Drive's `archive/` folder, not `inbox/`.
   - Supabase `master_catalog` should have the upserted rows (verify via admin app's database status metric).

If anything fails, the workflow run page shows full stack traces.

---

## Troubleshooting

**Apps Script fails with `GITHUB_PAT not found`** — Script Properties not set. Go back to Step 6.4.

**GitHub dispatch returns 404** — Wrong owner/repo or PAT lacks `repo` scope. Check Step 6.3 and Step 5.3.

**GitHub Action fails on `GDRIVE_SERVICE_ACCOUNT_JSON`** — Secret not set or contains wrong JSON. Re-paste the whole file contents.

**Action says "no files in inbox"** when there are files** — Service account doesn't have access. Re-share the folder (Step 3.3) and confirm the email is exact.

**Action runs but no rows upserted** — Check the action logs for "Engine output: 0 rows". The file might have a different structure than expected. Re-run the local dry-run scripts (`test_safilo_dryrun.py`) against the problem file to diagnose.

**Apps Script polls but nothing happens** — Check **Apps Script → Executions** (left sidebar). Should show recent runs of `pollInboxAndDispatch`. If failing, the log shows why.
