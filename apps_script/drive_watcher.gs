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
 * Install time-driven triggers for both pollers (every 5 minutes).
 * Idempotent — removes any existing matching triggers first.
 */
function setupTrigger() {
  const managed = new Set(["pollInboxAndDispatch", "pollGmailAndSaveAttachments"]);

  // Clear any existing triggers we manage
  ScriptApp.getProjectTriggers().forEach(t => {
    if (managed.has(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("pollGmailAndSaveAttachments")
    .timeBased()
    .everyMinutes(5)
    .create();

  ScriptApp.newTrigger("pollInboxAndDispatch")
    .timeBased()
    .everyMinutes(5)
    .create();

  console.log("Installed 2 time-driven triggers (every 5 min): " +
    "pollGmailAndSaveAttachments, pollInboxAndDispatch. " +
    "Verify in Apps Script editor → Triggers (clock icon).");
}

/**
 * Convenience: run once manually to confirm credentials work without
 * having to wait for the cron.
 */
function testDispatchNow() {
  pollInboxAndDispatch();
}


// ============================================================
// GMAIL → DRIVE BRIDGE
// ============================================================

const GMAIL_CONFIG = {
  // Subject must contain this string (case-insensitive in Gmail search).
  SUBJECT_CONTAINS: "Availability Safilo File",

  // How far back to look on each poll. 7 days is plenty for a daily file
  // and lets us recover if the trigger is paused for a day or two.
  LOOKBACK: "7d",
};

/**
 * Scan Gmail for matching messages and save CSV attachments to the inbox folder.
 * Tracks processed message IDs in Script Properties so we never save the same
 * attachment twice — without modifying the email itself.
 *
 * If any files are saved, immediately fires pollInboxAndDispatch so the
 * GitHub Action runs without waiting for the next 5-min tick.
 */
function pollGmailAndSaveAttachments() {
  const PROP_NAME = "PROCESSED_MESSAGE_IDS";
  const props = PropertiesService.getScriptProperties();
  const processedRaw = props.getProperty(PROP_NAME) || "";
  const processed = new Set(processedRaw.split(",").filter(x => x));

  const query = `subject:"${GMAIL_CONFIG.SUBJECT_CONTAINS}" has:attachment newer_than:${GMAIL_CONFIG.LOOKBACK}`;
  const threads = GmailApp.search(query, 0, 50);

  if (threads.length === 0) {
    console.log(`No emails match query: ${query}`);
    return;
  }

  console.log(`Found ${threads.length} matching thread(s) in last ${GMAIL_CONFIG.LOOKBACK}.`);

  const inbox = DriveApp.getFolderById(CONFIG.INBOX_FOLDER_ID);
  let saved = 0;
  const updatedProcessed = new Set(processed);

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      const msgId = message.getId();
      if (processed.has(msgId)) {
        return; // already handled in a previous run
      }
      updatedProcessed.add(msgId);

      message.getAttachments().forEach(att => {
        const name = att.getName();
        if (name.toLowerCase().endsWith(".csv")) {
          inbox.createFile(att);
          console.log(`Saved attachment '${name}' (msg ${msgId})`);
          saved++;
        } else {
          console.log(`Skipped non-CSV attachment '${name}' (msg ${msgId})`);
        }
      });
    });
  });

  // Trim memory to last 500 message IDs so Script Properties don't grow forever
  let arr = Array.from(updatedProcessed);
  if (arr.length > 500) arr = arr.slice(-500);
  props.setProperty(PROP_NAME, arr.join(","));

  console.log(`Done. Saved ${saved} new CSV attachment(s).`);

  if (saved > 0) {
    console.log("Triggering inbox dispatch immediately…");
    pollInboxAndDispatch();
  }
}

/**
 * Convenience: run manually to test the Gmail polling without waiting for the trigger.
 */
function testGmailPollNow() {
  pollGmailAndSaveAttachments();
}
