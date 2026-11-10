/* global Office, Excel */
/*
  Core for Excel → Sheets add-in:
  - constants & config
  - DOM helpers
  - shared state
  - Cloudflare Access sign-in (checks + waits)
  - Sheets file list
  - active worksheet tracking + local "Excel changed" detection (no network)
  - commodity parsing + totals

  NETWORK RULES (keeps Google Sheets API usage low):
  - /api/sheets              : sign-in check and "Reload Sheets list" only
  - /api/receiverRecordStatus: ONLY when "Read data" is pressed (see workflow.js)
  - /api/updateMasterFromReceiver + /api/upsertReceiverRecord: ONLY on Send
  Tab switches, renames and cell edits are handled entirely inside Excel.
*/

const BACKEND = "https://cmx-api.sheetscmx.workers.dev";

// ----------------- CONSTANTS / CONFIG -----------------

// Header fields on the receiver template.
const HEADER_FIELDS = [
  { key: "pf",      cell: "B11", cId: "c_pf" },
  { key: "dr",      cell: "B12", cId: "c_dr" },
  { key: "carrier", cell: "B13", cId: "c_carrier" },
  { key: "po",      cell: "E11", cId: "c_po" },
  { key: "rcv",     cell: "E12", cId: "c_rcv" },
  { key: "ver",     cell: "E13", cId: "c_ver" }
];

// Line item layout (multiple rows): A=ITEM NO, B=COMMODITY, C=GROSS, D=TARE, F=COST
const ITEM_START_ROW = 18;   // first row for items
const MAX_LINE_ITEMS = 20;   // safety cap

// Regex to extract id from Excel tab name, e.g. "...-11-02-345"
const ID_RE = /-\s*(\d{1,2})-(\d{1,2})-(\d{3})$/;

const LAST_SHEET_STORAGE_KEY = "cmx.lastSheetId";

// ----------------- DOM HELPERS -----------------

const $ = id => document.getElementById(id);
const show = el => el && el.classList.remove("hidden");
const hide = el => el && el.classList.add("hidden");
const sleep_ = ms => new Promise(r => setTimeout(r, ms));

// kind: "info" (default) | "success" | "warn" | "error"
function setStatus(txt, kind) {
  const box = $("statusBox");
  if (!box) return;
  if (txt && txt.length) {
    box.className = "notice notice-" + (kind || "info");
    box.textContent = txt;
  } else {
    box.textContent = "";
    box.className = "notice notice-info hidden";
  }
}

// Robust button enable/disable
function setButtonEnabled(buttonEl, enabled) {
  if (!buttonEl) return;
  buttonEl.disabled = !enabled;
  if (enabled) {
    buttonEl.classList.remove("disabled");
    buttonEl.setAttribute("aria-disabled", "false");
  } else {
    buttonEl.classList.add("disabled");
    buttonEl.setAttribute("aria-disabled", "true");
  }
}

// Apply/remove locked visual state to a whole section
function setSectionLocked(sectionEl, locked) {
  if (!sectionEl) return;
  if (locked) sectionEl.classList.add("locked-section");
  else sectionEl.classList.remove("locked-section");
}

function setGlobalLoading(isLoading, message) {
  const overlay = $("globalLoadingOverlay");
  const msgEl = $("globalLoadingText");
  if (!overlay || !msgEl) return;

  if (isLoading) {
    if (message) msgEl.textContent = message;
    overlay.classList.remove("hidden");
  } else {
    overlay.classList.add("hidden");
  }
}

// ----------------- RUNTIME STATE -----------------

let lastAuthOk = false;
let sheetsList = [];
let selectedSheetId = null;

// Active Excel worksheet (name can change on rename; id never does)
let currentActiveWorksheet = null;
let currentActiveWorksheetId = null;

// Data from the last "Read data" press
let headerValues = { pf: "", dr: "", carrier: "", po: "", rcv: "", ver: "" };
let createLineItems = [];   // [{ row, item, com, gross, tare, cost }]
let lastRead = null;        // { sheetId, sheetName, fileId, signature }

// True when the Excel cells differ from what was last read (detected locally)
let excelChangedSinceRead = false;

// Whether this RR# + PO# already exists in RECEIVER RECORDS.
// true = update, false = new, null = not checked / unknown.
let receiverRecordExists = null;

let createInFlight = false;
let sendInFlight = false;

// ----------------- CLOUDFLARE ACCESS SIGN-IN -----------------

let authDialog = null;          // open sign-in dialog, if any
let _authCheckPromise = null;   // de-dupes concurrent checks
let _verifyRunning = false;
let _verifyAgain = false;

// state: "checking" | "signedOut" | "waiting" | "verifying" | "error"
function setAuthView(state, msg, kind) {
  const progress = $("authProgress");
  const statusEl = $("authStatus");
  const errEl = $("authError");
  const btn = $("cfSignInBtn");
  if (!progress || !statusEl || !errEl || !btn) return;

  hide(progress);
  hide(errEl);
  show(btn);
  setButtonEnabled(btn, true);
  btn.textContent = "Connect";

  if (state === "checking") {
    statusEl.textContent = msg || "Checking your connection…";
    show(progress);
    hide(btn);
  } else if (state === "waiting") {
    statusEl.textContent = msg || "Waiting for you to finish signing in…";
    show(progress);
    btn.textContent = "Open sign-in window again";
  } else if (state === "verifying") {
    statusEl.textContent = msg || "Verifying sign-in…";
    show(progress);
    setButtonEnabled(btn, false);
  } else if (state === "error") {
    errEl.className = "notice notice-" + (kind || "error");
    errEl.textContent = msg || "Sign-in failed.";
    show(errEl);
    btn.textContent = "Try again";
  } else if (msg) {
    // signedOut with an explanation (e.g. session expired)
    errEl.className = "notice notice-" + (kind || "warn");
    errEl.textContent = msg;
    show(errEl);
  }
}

function setConnBadge_(ok) {
  const b = $("connBadge");
  if (!b) return;
  b.textContent = ok ? "Connected" : "Not connected";
  if (ok) b.classList.add("ok");
  else b.classList.remove("ok");
}

// One request to the Worker. This is the same call the add-in has always used
// to test the Cloudflare Access session (cookie is sent via credentials: "include").
// Returns { ok: true, list } or { ok: false, status } (status 0 = no response / blocked).
async function fetchSheetsList_(timeoutMs) {
  const ctrl = (typeof AbortController !== "undefined") ? new AbortController() : null;
  const timer = ctrl ? setTimeout(() => ctrl.abort(), timeoutMs || 20000) : null;
  try {
    const opts = { credentials: "include" };
    if (ctrl) opts.signal = ctrl.signal;
    const res = await fetch(`${BACKEND}/api/sheets`, opts);
    if (res.status === 200) {
      const json = await res.json();
      return { ok: true, list: Array.isArray(json) ? json : [] };
    }
    return { ok: false, status: res.status };
  } catch (err) {
    // Not signed in (Access redirected the request), offline, or timed out.
    return { ok: false, status: 0, error: err };
  } finally {
    if (timer) clearTimeout(timer);
  }
}

// opts: { attempts = 1, delayMs = 1500, afterDialog = false }
// Resolves true when signed in.
function checkAuth(opts) {
  if (_authCheckPromise) return _authCheckPromise;
  _authCheckPromise = runAuthCheck_(opts || {}).finally(() => { _authCheckPromise = null; });
  return _authCheckPromise;
}

async function runAuthCheck_(opts) {
  const attempts = opts.attempts || 1;
  const delayMs = opts.delayMs || 1500;
  const afterDialog = !!opts.afterDialog;

  let result = null;
  for (let i = 1; i <= attempts; i++) {
    if (!lastAuthOk) {
      if (afterDialog) {
        setAuthView("verifying", i > 1 ? `Verifying sign-in… (try ${i} of ${attempts})` : "Verifying sign-in…");
      } else {
        setAuthView("checking");
      }
    }

    result = await fetchSheetsList_();
    if (result.ok) {
      await onAuthOk_(result.list);
      return true;
    }

    // A real server error (e.g. 500 from the Worker) will not fix itself by retrying.
    if (result.status && result.status !== 401 && result.status !== 403) break;

    if (i < attempts) await sleep_(delayMs);
  }

  onAuthFail_(result, afterDialog);
  return false;
}

async function onAuthOk_(list) {
  lastAuthOk = true;
  sheetsList = list || [];
  setConnBadge_(true);
  show($("mainContent"));
  hide($("auth-section"));
  populateSheetsSelect();

  if (selectedSheetId) setStatus("");
  else setStatus("Choose a Google Sheets file to continue.");

  await detectActiveWorksheet();
  refreshUI();
}

function onAuthFail_(result, afterDialog) {
  lastAuthOk = false;
  setConnBadge_(false);
  hide($("mainContent"));
  show($("auth-section"));
  setStatus("");

  const status = result ? result.status : 0;

  // Sign-in window is still open: keep waiting for it.
  if (authDialog) {
    setAuthView("waiting");
    return;
  }

  if (status && status !== 401 && status !== 403) {
    setAuthView(
      "error",
      `You are signed in, but the CMX server returned an error (${status}). Wait a minute and try again.`
    );
  } else if (afterDialog) {
    setAuthView(
      "error",
      "Sign-in was not confirmed. Click Try again and finish the email code step. " +
      "If it keeps failing, check your internet connection."
    );
  } else {
    setAuthView("signedOut");
  }
}

// Called when any API call shows the session is gone (401 or blocked request).
function handleAuthLost_(msg) {
  lastAuthOk = false;
  setConnBadge_(false);
  hide($("mainContent"));
  show($("auth-section"));
  setStatus("");
  setAuthView("signedOut", msg || "Your sign-in expired. Connect again to continue. Your Excel data is untouched.");
}

// After the sign-in dialog reports back: check, wait, check again (max 3 tries).
async function verifyAfterSignIn_() {
  if (_verifyRunning) {
    _verifyAgain = true;
    return;
  }
  _verifyRunning = true;
  try {
    do {
      _verifyAgain = false;
      const ok = await checkAuth({ attempts: 3, delayMs: 1500, afterDialog: true });
      if (ok) break;
    } while (_verifyAgain);
  } finally {
    _verifyRunning = false;
  }
}

function openCFAccessDialog() {
  // Already confirming a sign-in: don't open a second window.
  if (_verifyRunning) return;

  // Only one Office dialog may be open at a time.
  if (authDialog) {
    try { authDialog.close(); } catch (e) {}
    authDialog = null;
  }

  setAuthView("waiting", "Opening sign-in window…");

  Office.context.ui.displayDialogAsync(
    `${BACKEND}/auth-check`,
    { height: 60, width: 35, displayInIframe: false },
    (res) => {
      if (res.status === Office.AsyncResultStatus.Failed) {
        const code = res.error && res.error.code ? ` (code ${res.error.code})` : "";
        setAuthView(
          "error",
          `Could not open the sign-in window${code}. Close any other sign-in window and try again.`
        );
        return;
      }

      const dlg = res.value;
      authDialog = dlg;
      setAuthView("waiting");

      // The /auth-check page reports success once Cloudflare Access lets it load.
      dlg.addEventHandler(Office.EventType.DialogMessageReceived, () => {
        try { dlg.close(); } catch (e) {}
        if (authDialog === dlg) authDialog = null;
        verifyAfterSignIn_();
      });

      // Fires when the window is closed (by the user, or by the page itself after
      // Access completes) or fails to load. Either way, test the session now.
      dlg.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
        if (arg && arg.error === 12006 && authDialog === dlg) authDialog = null;
        verifyAfterSignIn_();
      });
    }
  );
}

// ----------------- SHEETS FILE LIST -----------------

async function loadSheets() {
  setStatus("Reloading Sheets list…");
  const result = await fetchSheetsList_();
  if (result.ok) {
    sheetsList = result.list;
    populateSheetsSelect();
    setStatus("Sheets list reloaded.", "success");
    refreshUI();
    return;
  }
  if (result.status === 401 || result.status === 403 || result.status === 0) {
    handleAuthLost_();
    return;
  }
  setStatus("Reload failed: server returned " + result.status + ".", "error");
}

function readStoredSheetId_() {
  try { return window.localStorage.getItem(LAST_SHEET_STORAGE_KEY) || null; } catch (e) { return null; }
}

function storeSheetId_(id) {
  try {
    if (id) window.localStorage.setItem(LAST_SHEET_STORAGE_KEY, id);
  } catch (e) {}
}

// Rebuilds the dropdown, keeping the current choice (or the one used last time).
function populateSheetsSelect() {
  const sel = $("modify_sheet_select");
  if (!sel) return;
  sel.innerHTML = "";

  const first = document.createElement("option");
  first.value = "";
  first.textContent = "— choose sheet file —";
  sel.appendChild(first);

  sheetsList.forEach(f => {
    const o = document.createElement("option");
    o.value = f.id;
    o.textContent = f.name;
    sel.appendChild(o);
  });

  const wanted = selectedSheetId || readStoredSheetId_();
  const exists = wanted && sheetsList.some(f => f.id === wanted);
  sel.value = exists ? wanted : "";
  selectedSheetId = exists ? wanted : null;
}

// No network here: choosing a file only records the choice.
function onSheetSelected() {
  selectedSheetId = $("modify_sheet_select").value || null;
  storeSheetId_(selectedSheetId);

  if (!selectedSheetId) {
    setStatus("Choose a Google Sheets file to continue.");
  } else if (lastRead && lastRead.fileId !== selectedSheetId) {
    setStatus("Sheets file changed. Read data again to check it against this file.", "warn");
  } else if (!lastRead) {
    setStatus("Ready. Read the receiver from Excel.");
  } else {
    setStatus("");
  }
  refreshUI();
}

// ----------------- ACTIVE WORKSHEET TRACKING (Excel only) -----------------

function applyActiveSheet_(id, name) {
  currentActiveWorksheetId = id || null;
  currentActiveWorksheet = name || "";
}

// Re-reads the active tab's name + id from Excel. No network.
async function detectActiveWorksheet() {
  try {
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      ws.load("name,id");
      await ctx.sync();
      applyActiveSheet_(ws.id, ws.name);
    });
  } catch (err) {
    console.warn("detectActiveWorksheet error", err);
  }
  refreshUI();
}

function lineValuesHaveAny_(vals) {
  return vals.some(v => v && String(v).trim().length);
}

// Reads the receiver cells from the ACTIVE worksheet. No network.
// signature = fingerprint of every value that gets sent, used to spot edits.
async function readExcelSnapshot_() {
  let snap = null;
  await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    ws.load("name,id");

    const headerRanges = {};
    HEADER_FIELDS.forEach(h => {
      const rng = ws.getRange(h.cell);
      rng.load("values");
      headerRanges[h.key] = rng;
    });

    const endRow = ITEM_START_ROW + MAX_LINE_ITEMS - 1;
    const lineRange = ws.getRange(`A${ITEM_START_ROW}:F${endRow}`);
    lineRange.load("values");

    await ctx.sync();

    const header = {};
    HEADER_FIELDS.forEach(h => {
      const rng = headerRanges[h.key];
      const v = rng && rng.values && rng.values[0] ? rng.values[0][0] : "";
      header[h.key] = v == null ? "" : String(v);
    });

    const lines = [];
    const values = lineRange.values || [];
    for (let i = 0; i < values.length; i++) {
      const row = values[i] || [];
      const itemVal  = row[0];
      const comVal   = row[1];
      const grossVal = row[2];
      const tareVal  = row[3];
      const costVal  = row[5]; // F column

      if (lineValuesHaveAny_([itemVal, comVal, grossVal, tareVal, costVal])) {
        lines.push({
          row:   ITEM_START_ROW + i,
          item:  itemVal  == null ? "" : String(itemVal),
          com:   comVal   == null ? "" : String(comVal),
          gross: grossVal == null ? "" : String(grossVal),
          tare:  tareVal  == null ? "" : String(tareVal),
          cost:  costVal  == null ? "" : String(costVal)
        });
      }
    }

    snap = {
      sheetId: ws.id,
      sheetName: ws.name || "",
      header,
      lines,
      signature: JSON.stringify([header, lines])
    };
  });
  return snap;
}

// Compares Excel with the last read. Runs on tab switch, cell edit, rename,
// Resync and when the pane gets focus. Excel only — never calls Google Sheets.
async function checkExcelForChanges_() {
  if (createInFlight || sendInFlight) return;

  if (!lastRead) {
    await detectActiveWorksheet();
    return;
  }

  try {
    const snap = await readExcelSnapshot_();
    applyActiveSheet_(snap.sheetId, snap.sheetName);
    if (snap.sheetId === lastRead.sheetId) {
      excelChangedSinceRead = snap.signature !== lastRead.signature;
    }
  } catch (err) {
    console.warn("checkExcelForChanges error", err);
  }
  refreshUI();
}

// ----------------- NAME / DATE HELPERS -----------------

function isWorkspaceNameFormatValid(name) {
  if (!name) return false;
  return ID_RE.test(String(name).trim());
}

// Excel hands dates over as serial numbers (e.g. 46336). Show them as MM/DD/YYYY.
// Display only — the raw value is what gets sent, exactly as before.
function formatDateForDisplay_(value) {
  const s = value == null ? "" : String(value).trim();
  if (!s) return "";
  const n = Number(s);
  if (!Number.isFinite(n) || n < 20000 || n > 80000) return s;
  const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(n) * 86400000);
  const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
  const dd = String(d.getUTCDate()).padStart(2, "0");
  return `${mm}/${dd}/${d.getUTCFullYear()}`;
}

// ----------------- COMMODITY PARSING -----------------

// Normalize commodity text, capture any parenthetical notes, and detect GENERATED prefix
function parseCommodityForNotes(raw) {
  // Used ONLY for Master logic.
  const result = {
    material: "",
    note: "",
    isGenerated: false,
    poOverride: "" // per-line PO number, e.g. "P012345"
  };
  if (raw == null) return result;

  let text = String(raw).trim();

  // GENERATED: prefix
  const genMatch = text.match(/^\s*generated\s*:\s*(.*)$/i);
  if (genMatch) {
    result.isGenerated = true;
    text = genMatch[1].trim();
  }

  // Pull (...) into noteParts (parentheses only)
  const noteParts = [];
  text = text.replace(/\(([^)]*)\)/g, (_, inner) => {
    if (inner && inner.trim()) noteParts.push(inner.trim());
    return " ";
  });

  // Look for "PO# PXXXXX" (with messy spacing) anywhere in the string.
  // - Allow "PO#P12345", "PO # P12345", etc.
  // - Require P + at least 4 digits so we don't grab random junk.
  const poMatch = text.match(/\bPO\s*#\s*(P\d{4,})\b/i);
  if (poMatch) {
    const pCode = poMatch[1];     // e.g. "P012345"
    const idx   = poMatch.index;

    result.poOverride = pCode.trim().toUpperCase();

    // For Master material: drop everything from the PO segment onward
    // so any receiver-specific words after the PO do NOT affect Master grouping.
    text = text.slice(0, idx).trim();
  }

  // Normalize remaining commodity as the Master "material"
  result.material = text.replace(/\s+/g, " ").trim();

  if (noteParts.length) {
    result.note = noteParts.join("; ");
  }

  return result;
}

// ----------------- LINE HELPERS / TOTALS -----------------

function lineHasAny_(li) {
  if (!li) return false;
  return !!(
    (li.item && li.item.trim().length) ||
    (li.com && li.com.trim().length) ||
    (li.gross && li.gross.trim().length) ||
    (li.tare && li.tare.trim().length) ||
    (li.cost && li.cost.trim().length)
  );
}

// allow "$2.38", "1,234.5", etc.
function parseNumberOrNull_(val) {
  if (val == null) return null;
  const s = String(val).replace(/[^0-9.\-]/g, "").trim();
  if (!s) return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

// Totals over VALID lines only. A line is valid when it has an item code,
// a gross weight, a cost, and a net weight above zero.
// Split-PO nuance: when TARE is blank, GROSS already is the net weight.
// "cost" here is the same number that is sent as the receiver's net total.
function computeTotals_() {
  const totals = { lines: 0, valid: 0, gross: 0, tare: 0, net: 0, cost: 0, skipped: [] };

  for (const li of createLineItems) {
    if (!lineHasAny_(li)) continue;
    totals.lines++;

    const itemCode = (li.item || "").trim();
    if (!itemCode) {
      totals.skipped.push({ row: li.row, reason: "no item code" });
      continue;
    }

    const gross = parseNumberOrNull_(li.gross);
    const tare  = parseNumberOrNull_(li.tare);  // optional
    const price = parseNumberOrNull_(li.cost);
    if (gross == null || price == null) {
      totals.skipped.push({ row: li.row, reason: gross == null ? "no gross weight" : "no cost" });
      continue;
    }

    const net = (tare == null) ? gross : (gross - tare);
    if (!Number.isFinite(net) || net <= 0) {
      totals.skipped.push({ row: li.row, reason: "net weight is zero or less" });
      continue;
    }

    totals.valid++;
    totals.gross += gross;
    totals.tare += (tare == null ? 0 : tare);
    totals.net += net;
    totals.cost += net * price;
  }

  return totals;
}

function formatWeight_(n) {
  try {
    return Number(n).toLocaleString(undefined, { maximumFractionDigits: 2 });
  } catch (_) {
    return String(n);
  }
}

function formatMoney_(n) {
  try {
    return Number(n).toLocaleString(undefined, { style: "currency", currency: "USD" });
  } catch (_) {
    return "$" + Number(n).toFixed(2);
  }
}
