/* global Office, Excel */
/*
  Core for Excel → Sheets add-in:
  - constants & config
  - DOM helpers
  - shared state
  - auth / Sheets helpers
  - active worksheet tracking
  - workspace state helpers
  - line-items helpers
  - build/write helpers
  - basic "has data" predicates
*/

const BACKEND = "https://cmx-api.sheetscmx.workers.dev";

// ----------------- CONSTANTS / CONFIG -----------------

// Header fields shared by Create / Modify.
const HEADER_FIELDS = [
  { key: "pf",  cell: "B11", cId: "c_pf",  mId: "m_pf"  },
  { key: "dr",  cell: "B12", cId: "c_dr",  mId: "m_dr"  },
  { key: "carrier", cell: "B13", cId: "c_carrier", mId: "m_carrier" },
  { key: "po",  cell: "E11", cId: "c_po",  mId: "m_po"  },
  { key: "rcv", cell: "E12", cId: "c_rcv", mId: "m_rcv" },
  { key: "ver", cell: "E13", cId: "c_ver", mId: "m_ver" }
];

// Line item layout (multiple rows).
const ITEM_START_ROW = 18;   // first row for items
const MAX_LINE_ITEMS = 20;   // safety cap
const ITEM_COLUMNS = {
  item:  "A",    // ITEM NO
  com:   "B",    // COMMODITY
  gross: "C",    // GROSS
  tare:  "D",    // TARE
  cost:  "F",    // COST
  notes: "O"     // MATERIAL NOTES
};
const LINE_KEYS = ["item", "com", "gross", "tare", "cost"];

// Regex to extract id from Excel tab name, e.g. "...-11-02-345"
const ID_RE = /-\s*(\d{1,2})-(\d{1,2})-(\d{3})$/;

// ----------------- DOM HELPERS -----------------

const $ = id => document.getElementById(id);
const show = el => el && el.classList.remove("hidden");
const hide = el => el && el.classList.add("hidden");

function setStatus(txt) {
  const box = $("statusBox");
  if (!box) return;
  if (txt && txt.length) {
    box.textContent = txt;
    show(box);
  } else {
    box.textContent = "";
    hide(box);
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

let pollHandle = null;

let lastAuthOk = false;
let sheetsList = [];
let selectedSheetId = null;
let currentTabs = [];   // tab titles for selected sheet

let currentActiveWorksheet = null;
let lastDetectedWorksheet = null;

// Which sheet did we read fields from last?
let lastReadCreateSheetName = null;

let workspaceState = {
  type: "unknown",      // "existing" | "new-valid" | "title-invalid" | "unknown"
  identifier: null,     // -MM-DD-XXX extracted from Excel
  matchedTab: null      // tab title in Sheets (if any)
};

// Whether this RR# + PO# already exists in RECEIVER RECORDS (determines Create vs Modify)
let receiverRecordExists = null;
// Throttle expensive server checks (prevents Sheets API read quota bursts)
const RECEIVER_STATUS_TTL_MS = 30 * 1000;          // cache /api/receiverRecordStatus result
const WORKSPACE_EVAL_MIN_INTERVAL_MS = 10 * 1000;  // don't re-evaluate every 2s poll unless inputs/tab changed
let _rrStatusCache = { key: null, at: 0, exists: null };
let _rrStatusInFlight = null;
let _lastWorkspaceEvalKey = null;
let _lastWorkspaceEvalAt = 0;

function invalidateReceiverRecordStatusCache_() {
  _rrStatusCache = { key: null, at: 0, exists: null };
  _rrStatusInFlight = null;
  _lastWorkspaceEvalAt = 0; // allow immediate re-eval
}


let createInFlight = false;
let sendInFlight = false;

// dynamic line items (arrays of { item, com, gross, tare, cost })
let createLineItems = [];

// ----------------- AUTH / SHEETS -----------------

function openAuthDialog() {
  setStatus("Opening sign-in window...");
  Office.context.ui.displayDialogAsync(
    `${BACKEND}/auth`,
    { height: 60, width: 35, displayInIframe: true },
    (res) => {
      if (res.status === Office.AsyncResultStatus.Failed) {
        setStatus("Failed to open sign-in window.");
        return;
      }
      const dlg = res.value;
      dlg.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        if (arg && arg.message) {
          if (arg.message === "auth-success" || arg.message === "success") {
            setStatus("Authentication completed. Click Verify Sign-In.");
          } else {
            setStatus("Auth message: " + arg.message);
          }
        }
        try { dlg.close(); } catch (e) {}
      });
    }
  );
}

async function checkAuth() {
  try {
    const res = await fetch(`${BACKEND}/api/sheets`, { credentials: "include" });
    if (res.status === 200) {
      lastAuthOk = true;
      const json = await res.json();
      sheetsList = json || [];
      show($("mainContent"));
      hide($("auth-section"));
      populateSheetsSelect();
      $("authStatus").textContent = "Connected.";
      setStatus("Select a Google Sheets file to continue.");
      await detectActiveWorksheet();
      refreshSendButtonsState();
    } else {
      // 401 — not authenticated yet, show connect button
      lastAuthOk = false;
      hide($("mainContent"));
      show($("auth-section"));
      $("authStatus").textContent = "";
      setStatus("Click Connect to sign in with your company email.");
    }
  } catch (err) {
    lastAuthOk = false;
    hide($("mainContent"));
    show($("auth-section"));
    setStatus("Connection failed: " + (err.message || err));
  }
}

function openCFAccessDialog() {
  setStatus("Opening sign-in window...");
  $("authStatus").textContent = "Waiting for sign-in...";

  Office.context.ui.displayDialogAsync(
    `${BACKEND}/auth-check`,
    { height: 60, width: 35, displayInIframe: false },
    (res) => {
      if (res.status === Office.AsyncResultStatus.Failed) {
        setStatus("Failed to open sign-in window.");
        return;
      }
      const dlg = res.value;

      // Listen for message from the page (fires if messageParent works)
      dlg.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
        try { dlg.close(); } catch (e) {}
        setStatus("Verifying connection...");
        await checkAuth();
      });

      // Also listen for dialog close event — this fires when the user
      // closes the window manually OR after CF Access completes and
      // redirects back. Either way, attempt auth at this point.
      dlg.addEventHandler(Office.EventType.DialogEventReceived, async (arg) => {
        // arg.error === 12006 means the dialog was closed
        setStatus("Verifying connection...");
        await checkAuth();
      });
    }
  );
}

async function loadSheets() {
  setStatus("Reloading Sheets list...");
  try {
    const res = await fetch(`${BACKEND}/api/sheets`, { credentials: "include" });
    if (res.status === 401) {
      lastAuthOk = false;
      hide($("mainContent"));
      show($("auth-section"));
      setStatus("Session expired. Please reconnect.");
      return;
    }
    if (!res.ok) throw new Error("Failed to load sheets: " + res.status);
    const json = await res.json();
    sheetsList = json || [];
    populateSheetsSelect();
    setStatus("Sheets list reloaded.");
  } catch (err) {
    setStatus("Reload failed: " + (err.message || err));
  }
}

function populateSheetsSelect() {
  const sel = $("modify_sheet_select");
  if (!sel) return;
  sel.innerHTML = "<option value=''>— choose sheet file —</option>";
  sheetsList.forEach(f => {
    const o = document.createElement("option");
    o.value = f.id;
    o.textContent = f.name;
    sel.appendChild(o);
  });
}

async function onSheetSelected() {
  selectedSheetId = $("modify_sheet_select").value || null;
  currentTabs = [];
  workspaceState = { type: "unknown", identifier: null, matchedTab: null };

  if (!selectedSheetId) {
    setStatus("Select a Google Sheets file to continue.");
    setSectionLocked($("createPanel"), true);
    refreshSendButtonsState();
    return;
  }

  setStatus("Loading tabs for the selected file...");
  try {
    const res = await fetch(
      `${BACKEND}/api/tabs?sheetId=${encodeURIComponent(selectedSheetId)}`,
      { credentials: "include" }
    );
    if (!res.ok) {
      setStatus("Failed to load tabs: " + res.status);
      refreshSendButtonsState();
      return;
    }
    const data = await res.json();
    currentTabs = (data || []).map(t =>
      (t.properties && t.properties.title) ? t.properties.title : (t.title || t)
    );

    await evaluateWorkspaceState();
    updateModeBanner();
    updateResyncWarning();

    if (workspaceState.type === "existing" && workspaceState.matchedTab) {
      setStatus(
        `Workspace for this Excel tab already exists as "${workspaceState.matchedTab}". ` +
        "Sending will update that tab."
      );
    } else if (workspaceState.type === "title-invalid") {
      // evaluateWorkspaceState already set a detailed message about title format.
    } else {
      setStatus("Tabs loaded. Ready to read from Excel.");
    }

    // Section is now unlocked (but buttons still controlled by refreshSendButtonsState)
    setSectionLocked($("createPanel"), false);
  } catch (err) {
    console.error(err);
    setStatus("Error loading tabs: " + (err.message || err));
  } finally {
    refreshSendButtonsState();
  }
}

// ----------------- ACTIVE WORKSHEET TRACKING -----------------

async function detectActiveWorksheet(forceOverlay = false) {
  let showLoader = false;
  try {
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      ws.load("name");
      await ctx.sync();

      const currentName = ws.name || "";
      const tabChanged = currentName !== lastDetectedWorksheet;
      if (forceOverlay || tabChanged) {
        setGlobalLoading(true, "Updating for active Excel tab...");
        showLoader = true;
      }

      currentActiveWorksheet = currentName;
      lastDetectedWorksheet = currentName;

      const nameSpan = $("activeSheetName");
      if (nameSpan) nameSpan.textContent = currentName;

      // Always mirror active Excel tab into the workspace-name field
      updateCreateWorkspaceNameFromActive(currentName);

      // Re-evaluate workspace existence and warnings on each poll
      // Re-evaluate sparingly (prevents backend/Snapshot reminders from hammering Sheets)
      const rrVal = ($("c_rcv")?.value || "").trim();
      const poVal = ($("c_po")?.value || "").trim();
      const dtVal = ($("c_dr")?.value || "").trim();
      const evalKey = `${currentName}|${selectedSheetId || ""}|${rrVal}|${poVal}|${dtVal}`;
      const now = Date.now();

      const keyChanged = evalKey !== _lastWorkspaceEvalKey;
      const timeOk = (now - _lastWorkspaceEvalAt) >= WORKSPACE_EVAL_MIN_INTERVAL_MS;

      if (forceOverlay || tabChanged || keyChanged || timeOk) {
        _lastWorkspaceEvalKey = evalKey;
        _lastWorkspaceEvalAt = now;
        await evaluateWorkspaceState();
      }

      updateModeBanner();
      updateResyncWarning();
      refreshSendButtonsState();
    });
  } catch (err) {
    console.warn("detectActiveWorksheet error", err);
  } finally {
    if (showLoader) setGlobalLoading(false);
  }
}

// Update create workspace name from active Excel tab (always overwrite)
function updateCreateWorkspaceNameFromActive(activeName) {
  const input = $("create_workspace_name");
  if (input && activeName) {
    input.value = activeName;
  }
}

// ----------------- WORKSPACE STATE / WARNINGS -----------------

async function evaluateWorkspaceState() {
  const activeName =
    currentActiveWorksheet ||
    ($("activeSheetName") ? $("activeSheetName").textContent : "");

  const identifier = extractIdentifierFromName(activeName);
  workspaceState = { type: "unknown", identifier: identifier || null, matchedTab: null };

  // Prefer the actual Receiver / PO values (these are what get written to RECEIVER RECORDS)
  const rrFromUI = ($("c_rcv") && $("c_rcv").value ? $("c_rcv").value : "").trim();
  const poFromUI = ($("c_po") && $("c_po").value ? $("c_po").value : "").trim();
  const dateFromUI = ($("c_dr") && $("c_dr").value ? $("c_dr").value : "").trim();

  if (!activeName) {
    workspaceState.type = "unknown";
    return;
  }

  // If we have RR + PO + a selected sheet, we can *directly* determine Create vs Modify
  // by checking RECEIVER RECORDS (this is the source of truth).
  if (selectedSheetId && rrFromUI && poFromUI) {
    const key = `${selectedSheetId}|${rrFromUI}|${poFromUI}|${dateFromUI || ""}`;
    const now = Date.now();

    // Cache + de-dupe in-flight requests to avoid hammering the backend (which then hammers Sheets)
    if (_rrStatusCache.key === key && (now - _rrStatusCache.at) < RECEIVER_STATUS_TTL_MS && _rrStatusCache.exists !== null) {
      receiverRecordExists = _rrStatusCache.exists;
    } else {
      try {
        if (_rrStatusInFlight && _rrStatusInFlight.key === key) {
          receiverRecordExists = await _rrStatusInFlight.promise;
        } else {
          const qs = new URLSearchParams({
            sheetId: selectedSheetId,
            rrNumber: rrFromUI,
            poNumber: poFromUI
          });
          if (dateFromUI) qs.set("dateReceived", dateFromUI);

          _rrStatusInFlight = {
            key,
            promise: (async () => {
              const stResp = await fetch(`${BACKEND}/api/receiverRecordStatus?${qs.toString()}`, { credentials: "include" });
              const stJson = await stResp.json();
              return !!stJson.exists;
            })()
          };

          receiverRecordExists = await _rrStatusInFlight.promise;
        }

        _rrStatusCache = { key, at: Date.now(), exists: receiverRecordExists };
      } catch (e) {
        // If the status check fails, don't block the UI; just treat as "not uploaded"
        receiverRecordExists = false;
        _rrStatusCache = { key: null, at: 0, exists: null };
      } finally {
        if (_rrStatusInFlight && _rrStatusInFlight.key === key) _rrStatusInFlight = null;
      }
    }

    // We have enough information to set the mode.
    workspaceState.type = receiverRecordExists ? "existing" : "new-valid";
    workspaceState.identifier = identifier || null;
    workspaceState.matchedTab = identifier ? findTabByIdentifier(identifier) : null;
    return;
  }

  // If we don't have RR+PO yet, we cannot accurately check RECEIVER RECORDS.
  // Title identifier is just a hint/guardrail until the receiver is read.
  if (!identifier) {
    workspaceState.type = "title-invalid";
    workspaceState.identifier = null;

    if (selectedSheetId) {
      setStatus(
        "Active Excel tab name does not contain -MM-DD-XXX identifier (e.g. NOVA -1-17-025 or NOVA - 1-17-025). " +
        "Read the receiver to fill Receiver # and PO#, or rename the tab. " +
        "Send is disabled until Receiver # + PO# are present."
      );
    }
    return;
  }

  // We have an identifier but still lack RR+PO, so we can't check RECEIVER RECORDS.
  workspaceState.type = "needs-read";
  workspaceState.identifier = identifier;
  workspaceState.matchedTab = findTabByIdentifier(identifier);
}


// extract identifier -MM-DD-XXX from name
function extractIdentifierFromName(name) {
  if (!name || typeof name !== "string") return null;
  const m = name.match(ID_RE);
  if (!m) return null;
  return `-${m[1]}-${m[2]}-${m[3]}`;
}

// Attempt to find a tab whose title contains the identifier (in a few padded forms)
function findTabByIdentifier(identifier) {
  if (!identifier || !currentTabs || !currentTabs.length) return null;
  const parts = identifier.split("-");
  // parts: ["", MM, DD, XXX]
  const mm = parts[1];
  const dd = parts[2];
  const xxx = parts[3];

  const mmP = mm.padStart(2, "0");
  const ddP = dd.padStart(2, "0");

  const candidates = new Set();
  candidates.add(identifier);                // -M-D-XXX
  candidates.add(`-${mmP}-${dd}-${xxx}`);
  candidates.add(`-${mm}-${ddP}-${xxx}`);
  candidates.add(`-${mmP}-${ddP}-${xxx}`);
  candidates.add(`${mm}-${dd}-${xxx}`);
  candidates.add(`${mmP}-${ddP}-${xxx}`);

  for (const t of currentTabs) {
    const lower = t.toLowerCase();
    for (const idf of candidates) {
      if (!idf) continue;
      if (lower.includes(idf.toLowerCase())) return t;
    }
  }
  return null;
}

// Normalize commodity text, capture any parenthetical notes, and detect GENERATED prefix
function parseCommodityForNotes(raw) {
  // Used ONLY for Master logic (and optional auto-notes),
  // NOT for the 1:1 receiver copy COM text.
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
    const fullMatch = poMatch[0];     // e.g. "PO# P012345"
    const pCode     = poMatch[1];     // e.g. "P012345"
    const idx       = poMatch.index;

    result.poOverride = pCode.trim().toUpperCase();

    // For Master material: drop everything from the PO segment onward
    // so any receiver-specific words after the PO do NOT affect Master grouping.
    text = text.slice(0, idx).trim();
    // (The full raw text still goes to the receiver copy unchanged,
    //  via buildValuesMapFromUI.)
  }

  // Normalize remaining commodity as the Master "material"
  result.material = text.replace(/\s+/g, " ").trim();

  if (noteParts.length) {
    result.note = noteParts.join("; ");
  }

  return result;
}


function updateModeBanner() {
  const banner = $("modeBanner");
  if (!banner) return;

  const activeName =
    currentActiveWorksheet ||
    ($("activeSheetName") ? $("activeSheetName").textContent : "");

  // If we don't know the Excel tab yet or no sheet selected, hide
  if (!activeName || !selectedSheetId) {
    banner.className = "mode-banner hidden";
    banner.querySelector(".text").textContent = "Waiting for Excel tab...";
    return;
  }

  const labelEl = banner.querySelector(".label");
  const textEl = banner.querySelector(".text");

  if (!labelEl || !textEl) return;

  if (workspaceState.type === "title-invalid") {
    banner.className = "mode-banner invalid";
    labelEl.textContent = "Mode";
    textEl.textContent =
      "Cannot determine mode yet. Fill Receiver # and PO# (or rename the tab to include -MM-DD-XXX).";
  } else if (workspaceState.type === "existing") {
    banner.className = "mode-banner modify";
    labelEl.textContent = "Modifying";
    textEl.textContent =
      'Receiver already exists in "RECEIVER RECORDS" (Receiver # + month/day). Will update Master Receivers and refresh Date Uploaded.';
  } else if (workspaceState.type === "new-valid") {
    banner.className = "mode-banner create";
    labelEl.textContent = "Creating";
    textEl.textContent =
      'Receiver not found in "RECEIVER RECORDS" (Receiver # + PO#). Will create a new upload record, then update Master Receivers.';
  } else if (workspaceState.type === "needs-read") {
    banner.className = "mode-banner";
    labelEl.textContent = "Mode";
    textEl.textContent =
      "Read the receiver (Receiver # + PO#) to determine Creating vs Modifying.";
  } else {
    banner.className = "mode-banner";
    labelEl.textContent = "Mode";
    textEl.textContent = "Waiting for workspace status…";
  }
}

// Warning about “data read from a different sheet”
function updateResyncWarning() {
  const warningEl = $("sheetChangeWarning");
  if (!warningEl) return;

  const currentName = currentActiveWorksheet ||
    ($("activeSheetName") ? $("activeSheetName").textContent : "");

  let showWarning = false;
  let msg = "";

  if (hasAnyCreateData() &&
      lastReadCreateSheetName &&
      lastReadCreateSheetName !== currentName) {
    showWarning = true;
    msg = `Excel sheet changed. Click "Read Data From Excel" again before sending.`;
  }


  if (showWarning) {
    warningEl.textContent = msg;
    show(warningEl);
  } else {
    warningEl.textContent = "";
    hide(warningEl);
  }
}

// ----------------- MULTI-ROW LINE ITEMS HELPERS -----------------

function addLineItem() {
  createLineItems.push({ item: "", com: "", gross: "", tare: "", cost: "" });
  renderLineItems();
  updateNetTotalDisplay();
  refreshSendButtonsState();
}

// Derived net weight for UI (handles split-PO receivers where tare is blank)
function computeNetWeightForLine_(li) {
  const g = (li && li.gross != null) ? String(li.gross).trim() : "";
  const t = (li && li.tare != null) ? String(li.tare).trim() : "";

  const gNum = Number(g);
  const tNum = Number(t);
  if (g === "" || isNaN(gNum)) return "";

  // If tare is blank, treat gross as the net weight (common on split-PO rows)
  if (t === "" || isNaN(tNum)) return gNum;

  return gNum - tNum;
}

function updateNetTotalDisplay() {
  const valEl = $("netTotalValue");
  if (!valEl) return;

  let total = 0;
  createLineItems.forEach(li => {
    const net = computeNetWeightForLine_(li);
    const c = (li && li.cost != null) ? String(li.cost).trim() : "";
    const costNum = Number(c);
    if (net === "" || isNaN(costNum)) return;
    total += Number(net) * costNum;
  });

  try {
    valEl.textContent = total.toLocaleString(undefined, { style: "currency", currency: "USD" });
  } catch (_) {
    valEl.textContent = "$" + total.toFixed(2);
  }
}

function renderLineItems() {
  const container = $("createLineItemsContainer");
  const items = createLineItems;

  if (!container) return;
  container.innerHTML = "";

  if (!items.length) {
    const p = document.createElement("p");
    p.className = "status-msg";
    p.textContent = "No line items detected. Use \"Add Line\" to add one.";
    container.appendChild(p);
    return;
  }

  items.forEach((li, idx) => {
    const row = document.createElement("div");
    row.className = "line-item-row";
    row.dataset.index = String(idx);

    LINE_KEYS.forEach(key => {
      const input = document.createElement("input");
      input.value = li[key] || "";
      input.classList.add(`li-${key}`);

      if (key === "gross" || key === "tare" || key === "cost") {
        input.classList.add("li-num");
      }

      input.addEventListener("input", () => {
        li[key] = input.value;
        // Update derived fields in real-time
        if (key === "gross" || key === "tare" || key === "cost") {
          // Net is derived from gross/tare; recompute if either changes
          if (netInput) {
            const nw = computeNetWeightForLine_(li);
            netInput.value = (nw === "" ? "" : String(nw));
          }
          updateNetTotalDisplay();
        }
        refreshSendButtonsState();
      });

      // Insert NET column after TARE
      if (key === "tare") {
        row.appendChild(input);

        // Derived NET field (readonly)
        var netInput = document.createElement("input");
        const nw = computeNetWeightForLine_(li);
        netInput.value = (nw === "" ? "" : String(nw));
        netInput.readOnly = true;
        netInput.classList.add("li-net");
        row.appendChild(netInput);

        return; // continue to next key
      }

      row.appendChild(input);
    });

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "secondary-btn small-btn";
    removeBtn.textContent = "×";
    removeBtn.addEventListener("click", () => {
      items.splice(idx, 1);
      renderLineItems();
      updateNetTotalDisplay();
      refreshSendButtonsState();
    });
    row.appendChild(removeBtn);

    container.appendChild(row);
  });

  updateNetTotalDisplay();
}

// ----------------- BUILD VALUES MAP FROM UI -----------------

function buildValuesMapFromUI() {
  const valuesMap = {};

  // headers
  HEADER_FIELDS.forEach(h => {
    const el = $(h.cId);
    valuesMap[h.cell] = el ? el.value || "" : "";
  });

  // line items
  createLineItems.forEach((li, idx) => {
    const row = ITEM_START_ROW + idx;
    const any =
      (li.item && li.item.trim().length) ||
      (li.com && li.com.trim().length) ||
      (li.gross && li.gross.trim().length) ||
      (li.tare && li.tare.trim().length) ||
      (li.cost && li.cost.trim().length);

    if (!any) return;

    const rawCom = li.com || "";
    const parsed = parseCommodityForNotes(rawCom);

    // RECEIVER COPY: must be 1:1 with Excel input.
    valuesMap[`${ITEM_COLUMNS.item}${row}`]  = li.item  || "";
    valuesMap[`${ITEM_COLUMNS.com}${row}`]   = rawCom;         // keep exact text
    valuesMap[`${ITEM_COLUMNS.gross}${row}`] = li.gross || "";
    valuesMap[`${ITEM_COLUMNS.tare}${row}`]  = li.tare  || "";
    valuesMap[`${ITEM_COLUMNS.cost}${row}`]  = li.cost  || "";

    // Optional: still use parsed note for MATERIAL NOTES column
    if (parsed.note) {
      valuesMap[`${ITEM_COLUMNS.notes}${row}`] = parsed.note;
    }
  });

  return valuesMap;
}


// ----------------- WRITE TO SHEETS (BATCH + FALLBACK) -----------------

async function writeFieldsToTab(sheetId, tabName, valuesMap) {
  const batch = Object.entries(valuesMap).map(([cell, value]) => ({ cell, value }));

  // try batch first
  try {
    const res = await fetch(`${BACKEND}/api/writeBatch`, {
      method: "POST",
      credentials: "include",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sheetId, tabName, values: batch })
    });
    if (res.ok) return;

    // fallback to individual calls
    for (const it of batch) {
      const r = await fetch(`${BACKEND}/api/writeCell`, {
        method: "POST",
        credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          sheetId,
          tabName,
          cell: it.cell,
          value: it.value
        })
      });
      if (!r.ok) throw new Error("writeCell failed: " + r.status);
    }
  } catch (err) {
    throw err;
  }
}

// ----------------- “HAS DATA” HELPERS -----------------

function hasAnyHeaderValues() {
  for (const h of HEADER_FIELDS) {
    const el = $(h.cId);
    if (el && String(el.value || "").trim().length) return true;
  }
  return false;
}

function hasAnyLineItems() {
  return createLineItems.some(li => {
    return (
      (li.item && li.item.trim().length) ||
      (li.com && li.com.trim().length) ||
      (li.gross && li.gross.trim().length) ||
      (li.tare && li.tare.trim().length) ||
      (li.cost && li.cost.trim().length)
    );
  });
}

function hasAnyCreateData() {
  return hasAnyHeaderValues() || hasAnyLineItems();
}