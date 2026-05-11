/**
 * Apps Script — watches the Safilo "inbox" Drive folder and triggers the
 * GitHub Action when new files arrive.
 *
 * SETUP (do once after pasting):
 *   1. In Apps Script editor: Project Settings (⚙) → check
 *      "Show appsscript.json manifest file in editor".
 *   2. Fill the CONFIG block below with your real values.
 *   3. Save the project.
 *   4. Run `setupTrigger()` once (Run menu → setupTrigger). Authorize when prompted.
 *      This installs a time-driven trigger that polls the inbox every 5 minutes.
 *
 * After this, dropping any file into the inbox folder fires the GitHub
 * Action within ~5 minutes.
 */

// ============ CONFIG ============
const CONFIG = {
  // Drive folder ID where Safilo CSVs are dropped. Same ID you put in the
  // GDRIVE_INBOX_FOLDER_ID GitHub secret.
  INBOX_FOLDER_ID: "PASTE_INBOX_FOLDER_ID_HERE",

  // Your GitHub user/org and repo name.
  GITHUB_OWNER: "bgal-png",
  GITHUB_REPO: "Glasses-Import-Filler",

  // GitHub Personal Access Token with `repo` scope. Generate at:
  //   github.com → Settings → Developer settings → Personal access tokens → Tokens (classic)
  // Stored as a Script Property — NEVER hard-code it here.
};
// ============ END CONFIG ============


/**
 * Read GitHub PAT from Script Properties (set via Project Settings →
 * Script Properties → Add). Property name: GITHUB_PAT.
 */
function _getPat() {
  const pat = PropertiesService.getScriptProperties().getProperty("GITHUB_PAT");
  if (!pat) {
    throw new Error("GITHUB_PAT not found in Script Properties. " +
      "Set it under Project Settings → Script Properties.");
  }
  return pat;
}

/**
 * Poll the inbox folder. If there are any files, fire a repository_dispatch
 * at the GitHub Action. Uses Script Properties to remember the last-seen
 * file count so we don't fire when the folder is empty.
 */
function pollInboxAndDispatch() {
  const folder = DriveApp.getFolderById(CONFIG.INBOX_FOLDER_ID);
  const files = folder.getFiles();

  let names = [];
  while (files.hasNext()) {
    names.push(files.next().getName());
  }

  if (names.length === 0) {
    console.log("Inbox empty — nothing to dispatch.");
    return;
  }

  console.log(`Inbox has ${names.length} file(s): ${names.join(", ")}`);
  console.log("Dispatching GitHub workflow…");

  const url = `https://api.github.com/repos/${CONFIG.GITHUB_OWNER}/${CONFIG.GITHUB_REPO}/dispatches`;
  const payload = {
    event_type: "safilo-ingest",
    client_payload: {
      triggered_by: "apps_script",
      file_count: names.length,
      file_names: names.slice(0, 20),
      timestamp: new Date().toISOString(),
    },
  };

  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: {
      "Accept": "application/vnd.github+json",
      "Authorization": `Bearer ${_getPat()}`,
      "X-GitHub-Api-Version": "2022-11-28",
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const code = resp.getResponseCode();
  if (code >= 200 && code < 300) {
    console.log(`Dispatched successfully (HTTP ${code}).`);
  } else {
    console.error(`Dispatch FAILED — HTTP ${code}: ${resp.getContentText()}`);
    throw new Error(`GitHub dispatch failed: HTTP ${code}`);
  }
}

/**
 * Install a time-driven trigger that polls every 5 minutes.
 * Idempotent — removes any existing pollInboxAndDispatch trigger first.
 */
function setupTrigger() {
  // Clear any existing triggers for this function
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "pollInboxAndDispatch") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("pollInboxAndDispatch")
    .timeBased()
    .everyMinutes(5)
    .create();

  console.log("Installed time-driven trigger (every 5 min). " +
    "Verify in Apps Script editor → Triggers (clock icon).");
}

/**
 * Convenience: run once manually to confirm credentials work without
 * having to wait for the cron.
 */
function testDispatchNow() {
  pollInboxAndDispatch();
}
