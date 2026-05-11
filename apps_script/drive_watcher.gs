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
  // Drive folder IDs — same values you put in the corresponding GitHub secrets.
  INBOX_FOLDER_ID: "PASTE_INBOX_FOLDER_ID_HERE",
  LUXOTTICA_INBOX_FOLDER_ID: "18KI6O_0Zkvyo38nyD-ajH_v3Qma-Bm9M",

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
 * Generic GitHub repository_dispatch trigger. Used by both Safilo and
 * Luxottica pollers.
 */
function _dispatchWorkflow(eventType, clientPayload) {
  const url = `https://api.github.com/repos/${CONFIG.GITHUB_OWNER}/${CONFIG.GITHUB_REPO}/dispatches`;
  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: {
      "Accept": "application/vnd.github+json",
      "Authorization": `Bearer ${_getPat()}`,
      "X-GitHub-Api-Version": "2022-11-28",
    },
    payload: JSON.stringify({ event_type: eventType, client_payload: clientPayload || {} }),
    muteHttpExceptions: true,
  });
  const code = resp.getResponseCode();
  if (code >= 200 && code < 300) {
    console.log(`Dispatched ${eventType} successfully (HTTP ${code}).`);
  } else {
    console.error(`Dispatch ${eventType} FAILED — HTTP ${code}: ${resp.getContentText()}`);
    throw new Error(`GitHub dispatch ${eventType} failed: HTTP ${code}`);
  }
}

/**
 * Poll the Safilo inbox folder. If there are any files, fire a
 * repository_dispatch at the safilo-ingest workflow.
 */
function pollInboxAndDispatch() {
  const folder = DriveApp.getFolderById(CONFIG.INBOX_FOLDER_ID);
  const files = folder.getFiles();

  let names = [];
  while (files.hasNext()) {
    names.push(files.next().getName());
  }

  if (names.length === 0) {
    console.log("Safilo inbox empty — nothing to dispatch.");
    return;
  }

  console.log(`Safilo inbox has ${names.length} file(s): ${names.join(", ")}`);
  _dispatchWorkflow("safilo-ingest", {
    triggered_by: "apps_script",
    file_count: names.length,
    file_names: names.slice(0, 20),
    timestamp: new Date().toISOString(),
  });
}

/**
 * Poll the Luxottica inbox folder. If there are any files, fire a
 * repository_dispatch at the luxottica-ingest workflow.
 */
function pollLuxotticaInboxAndDispatch() {
  const folder = DriveApp.getFolderById(CONFIG.LUXOTTICA_INBOX_FOLDER_ID);
  const files = folder.getFiles();

  let names = [];
  while (files.hasNext()) {
    names.push(files.next().getName());
  }

  if (names.length === 0) {
    console.log("Luxottica inbox empty — nothing to dispatch.");
    return;
  }

  console.log(`Luxottica inbox has ${names.length} file(s): ${names.join(", ")}`);
  _dispatchWorkflow("luxottica-ingest", {
    triggered_by: "apps_script",
    file_count: names.length,
    file_names: names.slice(0, 20),
    timestamp: new Date().toISOString(),
  });
}

/**
 * Install time-driven triggers for all four pollers (every 5 minutes).
 * Idempotent — removes any existing matching triggers first.
 */
function setupTrigger() {
  const managed = new Set([
    "pollInboxAndDispatch",
    "pollGmailAndSaveAttachments",
    "pollLuxotticaInboxAndDispatch",
    "pollGmailAndFetchLuxotticaZip",
  ]);

  // Clear any existing triggers we manage
  ScriptApp.getProjectTriggers().forEach(t => {
    if (managed.has(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Gmail watchers (find new emails, save files to Drive)
  ScriptApp.newTrigger("pollGmailAndSaveAttachments").timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger("pollGmailAndFetchLuxotticaZip").timeBased().everyMinutes(5).create();

  // Drive inbox watchers (fire GitHub Action when files are present)
  ScriptApp.newTrigger("pollInboxAndDispatch").timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger("pollLuxotticaInboxAndDispatch").timeBased().everyMinutes(5).create();

  console.log("Installed 4 time-driven triggers (every 5 min): " +
    "pollGmailAndSaveAttachments, pollGmailAndFetchLuxotticaZip, " +
    "pollInboxAndDispatch, pollLuxotticaInboxAndDispatch. " +
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


// ============================================================
// LUXOTTICA — GMAIL → AZURE BLOB FETCH → DRIVE
// ============================================================
//
// The Luxottica email doesn't contain the file as an attachment. Instead it
// has a "ZOBRAZENÍ" button linking to a time-limited Azure Blob URL with a
// SAS signature. We extract that URL from the HTML body, fetch the ZIP
// directly via UrlFetchApp, and drop the ZIP into Drive Luxottica/inbox.
// The GitHub Action unzips and processes the .xlsx inside.

const LUXOTTICA_GMAIL_CONFIG = {
  // Subject must contain this string (case-insensitive in Gmail search).
  SUBJECT_CONTAINS: "Požadavek databáze položek ze dne",
  LOOKBACK: "7d",
};

function pollGmailAndFetchLuxotticaZip() {
  const PROP_NAME = "LUXOTTICA_PROCESSED_MESSAGE_IDS";
  const props = PropertiesService.getScriptProperties();
  const processedRaw = props.getProperty(PROP_NAME) || "";
  const processed = new Set(processedRaw.split(",").filter(x => x));

  const query = `subject:"${LUXOTTICA_GMAIL_CONFIG.SUBJECT_CONTAINS}" newer_than:${LUXOTTICA_GMAIL_CONFIG.LOOKBACK}`;
  const threads = GmailApp.search(query, 0, 50);

  if (threads.length === 0) {
    console.log(`No Luxottica emails match query: ${query}`);
    return;
  }

  console.log(`Found ${threads.length} matching Luxottica thread(s) in last ${LUXOTTICA_GMAIL_CONFIG.LOOKBACK}.`);

  const inbox = DriveApp.getFolderById(CONFIG.LUXOTTICA_INBOX_FOLDER_ID);
  let saved = 0;
  const updated = new Set(processed);

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      const msgId = message.getId();
      if (processed.has(msgId)) {
        return; // already handled
      }
      updated.add(msgId);

      const body = message.getBody();
      // Match all Azure Blob URLs in the body, then pick the one that's actually
      // the ZIP download (path ends in .zip and has a SAS signature). The email
      // also embeds the company logo from an Azure URL — must not match that.
      const allMatches = body.match(/https:\/\/[^"'\s<>]+blob\.core\.windows\.net[^"'\s<>]+/g) || [];
      const candidates = allMatches
        .map(u => u.replace(/&amp;/g, "&"))
        .filter(u => /\.zip(\?|$)/i.test(u) && /[?&]sig=/i.test(u));
      if (candidates.length === 0) {
        console.log(`  msg ${msgId}: no ZIP download URL found among ${allMatches.length} Azure URL(s) — skipping`);
        return;
      }
      // If multiple, pick the longest (most-qualified — usually the one with
      // the full SAS token).
      candidates.sort((a, b) => b.length - a.length);
      const url = candidates[0];
      console.log(`  msg ${msgId}: found ZIP URL (length ${url.length}) among ${allMatches.length} blob URL(s)`);

      // Filename = last path segment of URL, before query string, URL-decoded.
      let filename;
      try {
        filename = decodeURIComponent(url.split("?")[0].split("/").pop());
      } catch (e) {
        filename = `Luxottica_${new Date().toISOString().replace(/[:.]/g, "-")}.zip`;
      }
      if (!filename.toLowerCase().endsWith(".zip")) {
        filename += ".zip";
      }

      try {
        const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        const code = resp.getResponseCode();
        if (code !== 200) {
          console.error(`  msg ${msgId}: HTTP ${code} fetching ZIP — skipping (link may have expired)`);
          return;
        }
        const blob = resp.getBlob().setName(filename);
        inbox.createFile(blob);
        console.log(`  msg ${msgId}: saved '${filename}' (${Math.round(resp.getContent().length / 1024)} KB)`);
        saved++;
      } catch (e) {
        console.error(`  msg ${msgId}: download failed: ${e}`);
      }
    });
  });

  let arr = Array.from(updated);
  if (arr.length > 500) arr = arr.slice(-500);
  props.setProperty(PROP_NAME, arr.join(","));

  console.log(`Done. Saved ${saved} new Luxottica ZIP(s).`);

  if (saved > 0) {
    console.log("Triggering Luxottica ingest workflow immediately…");
    pollLuxotticaInboxAndDispatch();
  }
}

/**
 * Convenience: run manually to test the Luxottica Gmail polling.
 */
function testLuxotticaPollNow() {
  pollGmailAndFetchLuxotticaZip();
}
