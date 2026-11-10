/* global Excel */
/*
  Workflow layer:
  - "Read data": read the receiver from Excel, then ONE check against RECEIVER RECORDS
  - build the payloads for the Worker (unchanged shapes)
  - send: update Master Receivers, then upsert RECEIVER RECORDS
  - refreshUI(): single place that draws the pane and decides whether Send is allowed
*/

let pendingSendPlan = null;

const CHECK_ICON_ =
  '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" ' +
  'stroke-linecap="round" stroke-linejoin="round" role="img" aria-label="Done"><path d="M20 6 9 17l-5-5"/></svg>';

// ----------------- READ FROM EXCEL -----------------

// The only place (besides Send) that talks to Google Sheets.
// Returns { exists: true|false|null, problem: string|null }.
async function fetchReceiverRecordStatus_(rrNumber, poNumber, dateReceived) {
  try {
    const qs = new URLSearchParams({
      sheetId: selectedSheetId,
      rrNumber,
      poNumber
    });
    if (dateReceived) qs.set("dateReceived", dateReceived);

    const res = await fetch(`${BACKEND}/api/receiverRecordStatus?${qs.toString()}`, { credentials: "include" });
    if (res.status === 401) return { exists: null, problem: "auth" };
    if (res.status === 404) return { exists: null, problem: "This Sheets file has no RECEIVER RECORDS tab." };
    if (!res.ok) return { exists: null, problem: "Could not check Receiver Records (" + res.status + ")." };

    const json = await res.json();
    return { exists: !!json.exists, problem: null };
  } catch (e) {
    return { exists: null, problem: "network" };
  }
}

async function readFieldsFromExcel_Create() {
  if (!selectedSheetId) {
    setStatus("Choose a Google Sheets file first.", "warn");
    return;
  }
  if (createInFlight || sendInFlight) return;

  createInFlight = true;
  refreshUI();
  setGlobalLoading(true, "Reading receiver from Excel…");

  try {
    // 1) Excel (local)
    const snap = await readExcelSnapshot_();

    // Always take the tab's CURRENT name + id from this read. This is what fixes
    // "renamed the tab, re-read, Send still greyed out".
    applyActiveSheet_(snap.sheetId, snap.sheetName);

    headerValues = snap.header;
    createLineItems = snap.lines;
    lastRead = {
      sheetId: snap.sheetId,
      sheetName: snap.sheetName,
      fileId: selectedSheetId,
      signature: snap.signature
    };
    excelChangedSinceRead = false;
    receiverRecordExists = null;

    // 2) Google Sheets (one request): new receiver or update?
    const rr = (headerValues.rcv || "").trim();
    const po = (headerValues.po || "").trim();
    const dr = (headerValues.dr || "").trim();

    let problem = null;
    if (rr && po) {
      setGlobalLoading(true, "Checking Receiver Records…");
      const st = await fetchReceiverRecordStatus_(rr, po, dr);
      receiverRecordExists = st.exists;
      problem = st.problem;
    }

    if (problem === "auth") {
      handleAuthLost_();
      return;
    }

    const n = createLineItems.length;
    const readMsg = `Read ${n} line${n === 1 ? "" : "s"} from "${snap.sheetName}".`;
    if (problem === "network") {
      setStatus(readMsg + " Could not reach the server to check Receiver Records; it will be checked when you send.", "warn");
    } else if (problem) {
      setStatus(readMsg + " " + problem, "warn");
    } else {
      setStatus(readMsg, "success");
    }
  } catch (err) {
    console.error(err);
    setStatus("Could not read Excel: " + (err.message || err) + " If a cell is being edited, press Enter and try again.", "error");
  } finally {
    createInFlight = false;
    setGlobalLoading(false);
    refreshUI();
  }
}

// ----------------- PAYLOADS -----------------

function buildReceiverRecordPayloadFromUI() {
  const purchaseFrom = String(headerValues.pf  || "").trim();
  const dateReceived = String(headerValues.dr  || "").trim();
  const poNumber     = String(headerValues.po  || "").trim();
  const rrNumber     = String(headerValues.rcv || "").trim();

  if (!rrNumber || !dateReceived) {
    return { payload: null, error: "Receiver # and Date Received are required to log RECEIVER RECORDS." };
  }

  // Net Total = sum of (NET * PRICE) for valid, coded line items (see computeTotals_).
  const totals = computeTotals_();

  // If we couldn't compute it reliably, send null and let the server write a live formula instead.
  const netTotalOut = totals.valid > 0 ? Number(totals.cost.toFixed(2)) : null;

  return {
    payload: {
      sheetId: selectedSheetId,
      rrNumber,
      poNumber,
      purchaseFrom,
      dateReceived,
      netTotal: netTotalOut
    },
    error: null
  };
}

function buildMasterPayloadFromUI() {
  const supplier = String(headerValues.pf      || "").trim();
  const date     = String(headerValues.dr      || "").trim();
  const carrier  = String(headerValues.carrier || "").trim();
  const poNumber = String(headerValues.po      || "").trim();
  const rrNumber = String(headerValues.rcv     || "").trim();

  // If there is no RR or PO, we can't meaningfully update Master.
  if (!rrNumber || !poNumber) {
    return { payload: null, error: null };
  }

  const lines = [];
  let generatedMissingCode = null;

  function parseNumber(val) {
    const n = parseNumberOrNull_(val);
    return n == null ? "" : n;
  }

  for (const li of createLineItems) {
    if (!lineHasAny_(li)) continue;

    const parsed   = parseCommodityForNotes(li.com);
    const itemCode = (li.item || "").trim();

    // Skip Master if no item code
    if (!itemCode) {
      if (parsed.isGenerated && !generatedMissingCode) {
        generatedMissingCode = parsed.material || "Generated line";
      }
      continue;
    }

    const gross = parseNumber(li.gross);
    const tare  = parseNumber(li.tare);

    // SPLIT-PO nuance:
    // When TARE is blank, users often enter NET directly in the GROSS cell.
    // Treat blank TARE as "gross already equals net".
    let net = "";
    if (gross !== "") {
      net = (tare === "") ? gross : (gross - tare);
      if (!Number.isFinite(net) || net <= 0) net = "";
    }
    const price = parseNumber(li.cost);

    // Per-line PO: use the inline PO# if present, otherwise fallback to header PO
    const effectivePo = parsed.poOverride || poNumber;

    lines.push({
      itemCode,
      material: parsed.material || "",
      materialNotes: parsed.note || "",
      net: net === "" ? "" : net,
      price: price === "" ? "" : price,
      extension: "",                 // let Sheets' formulas handle extensions
      poWeight: "",                  // updated/derived in Master
      linePoNumber: effectivePo      // per-line PO for multi-PO situations
    });
  }

  if (generatedMissingCode) {
    return { payload: null, error: `Generated line "${generatedMissingCode}" requires an item code.` };
  }

  if (!lines.length) {
    // No valid lines for Master (all missing item codes etc.)
    return { payload: null, error: null };
  }

  return {
    payload: {
      sheetId: selectedSheetId,
      rrNumber,
      date,
      supplier,
      status: "",           // Master formulas handle these
      datePaid: "",
      term: "",             // computed in Master from PO/Supplier if needed
      dueDate: "",
      daysTillDue: "",
      poNumber,             // header PO (used when no per-line override)
      poStatus: "",
      poSageClosed: "",
      receiverSageEntry: "",
      notes: "",
      carrier,
      lines
    },
    error: null
  };
}

function ensureGeneratedRowsHaveCodes() {
  for (const li of createLineItems) {
    if (!lineHasAny_(li)) continue;

    const parsed = parseCommodityForNotes(li.com);
    const itemCode = (li.item || "").trim();

    if (parsed.isGenerated && !itemCode) {
      return `Generated line "${parsed.material || ""}" (row ${li.row}) requires an item code.`;
    }
  }
  return null;
}

// ----------------- SEND GATE -----------------

// Returns the reason Send is not allowed, or null when it is. All local checks.
function getSendBlocker_() {
  if (!selectedSheetId) return "Choose a Google Sheets file.";
  if (!lastRead) return "Read the receiver from Excel.";

  if (lastRead.sheetId !== currentActiveWorksheetId) {
    return "This is a different Excel tab than the one that was read. Read data from this tab.";
  }
  if (excelChangedSinceRead) {
    return "Excel values changed since the last read. Read data again.";
  }
  if (lastRead.fileId !== selectedSheetId) {
    return "Sheets file changed since the last read. Read data again.";
  }

  if (!isWorkspaceNameFormatValid(currentActiveWorksheet)) {
    return "Rename the Excel tab so it ends with -MM-DD-XXX (e.g. NOVA -1-17-025).";
  }

  const rr = (headerValues.rcv || "").trim();
  const po = (headerValues.po || "").trim();
  const dr = (headerValues.dr || "").trim();
  if (!rr || !po) return "Receiver # (cell E12) and Our PO # (cell E11) are required.";
  if (!dr) return "Date received (cell B12) is required.";

  const generatedError = ensureGeneratedRowsHaveCodes();
  if (generatedError) return generatedError;

  return null;
}

function buildSendPlan() {
  const blocker = getSendBlocker_();
  if (blocker) return { ok: false, message: blocker };

  // Sending always:
  // 1) updates Master Receivers, and
  // 2) upserts a row in RECEIVER RECORDS.
  return { ok: true, action: "send", workspaceName: currentActiveWorksheet };
}

async function openSendReviewModal() {
  if (sendInFlight || createInFlight) return;

  // Last local check before sending: is Excel still what we read? (no network)
  try {
    const snap = await readExcelSnapshot_();
    applyActiveSheet_(snap.sheetId, snap.sheetName);
    if (lastRead && snap.sheetId === lastRead.sheetId) {
      excelChangedSinceRead = snap.signature !== lastRead.signature;
    }
  } catch (err) {
    console.warn("pre-send Excel check failed", err);
    setStatus("Could not check Excel before sending. If a cell is being edited, press Enter and try again.", "warn");
    refreshUI();
    return;
  }
  refreshUI();

  const plan = buildSendPlan();
  if (!plan.ok) {
    setStatus(plan.message || "Unable to send.", "warn");
    return;
  }

  pendingSendPlan = plan;
  const modal = $("actionModal");
  const title = $("actionModalTitle");
  const body = $("actionModalBody");

  const totals = computeTotals_();
  const rr = (headerValues.rcv || "").trim();
  const po = (headerValues.po || "").trim();

  if (title) title.textContent = "Send receiver to Google Sheets";

  if (body) {
    const action =
      receiverRecordExists === true
        ? "This receiver is already in RECEIVER RECORDS. Sending updates its row and its Master Receivers lines."
        : receiverRecordExists === false
          ? "This is a new receiver. Sending adds it to RECEIVER RECORDS and Master Receivers."
          : "Sending updates Master Receivers and adds or updates the RECEIVER RECORDS row.";

    body.textContent =
      `Receiver # ${rr} · PO ${po}\n` +
      `Total cost ${formatMoney_(totals.cost)} (${totals.valid} valid line${totals.valid === 1 ? "" : "s"})\n\n` +
      action;
  }

  if (modal) modal.classList.remove("hidden");
}

function closeActionModal() {
  const modal = $("actionModal");
  if (modal) modal.classList.add("hidden");
  pendingSendPlan = null;
}

// ----------------- SEND -----------------

async function confirmPendingSend() {
  if (!pendingSendPlan || sendInFlight) {
    closeActionModal();
    return;
  }

  closeActionModal();

  sendInFlight = true;
  refreshUI();
  setGlobalLoading(true, "Sending receiver to Google Sheets…");
  setStatus("Sending to Google Sheets…");

  let sawAuthProblem = false;
  let masterUpdated = false;

  try {
    const rrResult = buildReceiverRecordPayloadFromUI();
    if (rrResult.error) throw new Error(rrResult.error);

    const masterResult = buildMasterPayloadFromUI();
    if (masterResult.error) throw new Error(masterResult.error);

    // 1) Update Master Receivers
    let masterFailStatus = null;

    if (masterResult.payload) {
      setGlobalLoading(true, "Updating Master Receivers…");
      try {
        const mrRes = await fetch(`${BACKEND}/api/updateMasterFromReceiver`, {
          method: "POST",
          credentials: "include",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(masterResult.payload)
        });
        masterUpdated = mrRes.ok;
        if (!mrRes.ok) {
          masterFailStatus = mrRes.status;
          if (mrRes.status === 401) sawAuthProblem = true;
        }
      } catch (e) {
        masterUpdated = false;
        masterFailStatus = e.message || String(e);
        sawAuthProblem = true; // blocked request: may be an expired session
        console.warn("updateMasterFromReceiver error:", e);
      }
    }

    // 2) Upsert RECEIVER RECORDS row (server stamps Date Uploaded)
    setGlobalLoading(true, "Logging in Receiver Records…");
    let rrUpRes;
    try {
      rrUpRes = await fetch(`${BACKEND}/api/upsertReceiverRecord`, {
        method: "POST",
        credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(rrResult.payload)
      });
    } catch (e) {
      sawAuthProblem = true;
      throw new Error("Receiver Records update could not reach the server.");
    }
    if (!rrUpRes.ok) {
      if (rrUpRes.status === 401) sawAuthProblem = true;
      throw new Error("Receiver Records update failed: " + rrUpRes.status);
    }

    // The server tells us whether it added a row or updated one.
    let action = null;
    try {
      const upJson = await rrUpRes.json();
      action = upJson && upJson.action ? upJson.action : null;
    } catch (_) {}

    let modeText;
    if (action === "updated") modeText = "updated existing upload record";
    else if (action === "appended") modeText = "created new upload record";
    else modeText = receiverRecordExists ? "updated existing upload record" : "created new upload record";

    // From here on this receiver exists in RECEIVER RECORDS.
    receiverRecordExists = true;

    if (!masterResult.payload) {
      setStatus(`Done: ${modeText}. Master Receivers was not updated because no line has an item code.`, "warn");
    } else if (!masterUpdated) {
      const why = masterFailStatus === 429
        ? "another update is in progress. Wait a few seconds and send again."
        : `error ${masterFailStatus}. Send again to retry.`;
      setStatus(`Done: ${modeText}, but Master Receivers was NOT updated: ${why}`, "error");
    } else {
      setStatus(`Done: ${modeText} and updated Master Receivers.`, "success");
    }
  } catch (err) {
    console.error(err);
    const extra = masterUpdated ? " Master Receivers was already updated; send again to finish the Receiver Records log." : "";
    setStatus("Send failed: " + (err.message || err) + extra, "error");
  } finally {
    sendInFlight = false;
    setGlobalLoading(false);
    refreshUI();
  }

  // If a request was rejected or blocked, find out whether the sign-in expired.
  // (Only runs after a failure, so it adds no traffic to normal sends.)
  if (sawAuthProblem) {
    const check = await fetchSheetsList_();
    if (!check.ok && (check.status === 401 || check.status === 403 || check.status === 0)) {
      handleAuthLost_("Your sign-in expired, so the send did not finish. Connect again, then press Send again.");
    }
  }
}

// ----------------- DRAW THE PANE -----------------

function setStepBadge_(el, number, done) {
  if (!el) return;
  if (done) {
    el.classList.add("done");
    el.innerHTML = CHECK_ICON_;
  } else {
    el.classList.remove("done");
    el.textContent = String(number);
  }
}

function refreshUI() {
  const hasFile = !!selectedSheetId;
  const name = currentActiveWorksheet || "";
  const nameValid = isWorkspaceNameFormatValid(name);

  const onReadTab = !!lastRead && lastRead.sheetId === currentActiveWorksheetId;
  const fileSame = !!lastRead && lastRead.fileId === selectedSheetId;
  const fresh = onReadTab && fileSame && !excelChangedSinceRead;
  const busy = createInFlight || sendInFlight;

  // --- sections
  setSectionLocked($("tabPanel"), !hasFile);
  setSectionLocked($("createPanel"), !hasFile);

  // --- step 1
  setStepBadge_($("step1Badge"), 1, hasFile);

  // --- step 2: tab
  const nameEl = $("activeSheetName");
  if (nameEl) nameEl.textContent = name || "—";

  const nameWarn = $("tabNameWarning");
  if (nameWarn) {
    if (name && !nameValid) {
      nameWarn.textContent =
        "Tab name must end with -MM-DD-XXX (e.g. NOVA -1-17-025). Rename the tab, then press Resync if it does not update here.";
      show(nameWarn);
    } else {
      nameWarn.textContent = "";
      hide(nameWarn);
    }
  }

  // "Read data again" notice (never touches the Resync button)
  const staleEl = $("sheetChangeWarning");
  if (staleEl) {
    let msg = "";
    if (lastRead && !onReadTab) {
      msg = `The last read came from another tab ("${lastRead.sheetName}"). Read data from this tab.`;
    } else if (lastRead && excelChangedSinceRead) {
      msg = "Excel values changed since the last read. Read data again before sending.";
    } else if (lastRead && !fileSame) {
      msg = "Sheets file changed since the last read. Read data again.";
    }
    staleEl.textContent = msg;
    if (msg) show(staleEl);
    else hide(staleEl);
  }

  // receiver details (only for the tab they were read from)
  const detailsRow = $("detailsRow");
  const detailsList = $("detailsList");
  const toggle = $("detailsToggle");
  if (onReadTab) {
    const rr = (headerValues.rcv || "").trim();
    const po = (headerValues.po || "").trim();
    const pf = (headerValues.pf || "").trim();
    const dr = formatDateForDisplay_(headerValues.dr);

    const parts = [];
    if (rr) parts.push("RR " + rr);
    if (po) parts.push("PO " + po);
    if (pf) parts.push(pf);
    if (dr) parts.push(dr);
    const summary = $("detailsSummary");
    if (summary) summary.textContent = parts.length ? parts.join(" · ") : "No header values found on this tab.";

    const required = { rcv: true, po: true, dr: true };
    HEADER_FIELDS.forEach(h => {
      const el = $(h.cId);
      if (!el) return;
      const raw = (headerValues[h.key] || "").trim();
      const shown = h.key === "dr" ? formatDateForDisplay_(raw) : raw;
      if (shown) {
        el.textContent = shown;
        el.classList.remove("missing");
      } else {
        el.textContent = required[h.key] ? `Missing (${h.cell})` : "—";
        if (required[h.key]) el.classList.add("missing");
        else el.classList.remove("missing");
      }
    });

    show(detailsRow);
    if (toggle && toggle.getAttribute("aria-expanded") === "true") show(detailsList);
    else hide(detailsList);

    [detailsRow, detailsList].forEach(el => {
      if (!el) return;
      if (fresh) el.classList.remove("is-stale");
      else el.classList.add("is-stale");
    });
  } else {
    hide(detailsRow);
    hide(detailsList);
  }

  // read button
  const readBtn = $("create_prefill");
  if (readBtn) {
    readBtn.textContent = onReadTab ? "Read data again" : "Read data from Excel";
    if (fresh) readBtn.classList.remove("btn-attn");
    else readBtn.classList.add("btn-attn");
    setButtonEnabled(readBtn, hasFile && !busy);
  }

  setStepBadge_($("step2Badge"), 2, fresh);

  // --- step 3: totals + send
  const totalsBlock = $("totalsBlock");
  const totalsEmpty = $("totalsEmpty");
  const badge = $("modeBanner");
  const modeNote = $("modeNote");

  if (onReadTab) {
    const t = computeTotals_();
    $("grossTotalValue").textContent = formatWeight_(t.gross);
    $("tareTotalValue").textContent = formatWeight_(t.tare);
    $("netWeightTotalValue").textContent = formatWeight_(t.net);
    $("netTotalValue").textContent = formatMoney_(t.cost);

    const note = $("linesNote");
    if (note) {
      note.textContent = "";
      const head = document.createElement("span");
      head.textContent = t.lines === 0
        ? "No line items found in rows 18 to 37."
        : `${t.valid} of ${t.lines} line${t.lines === 1 ? "" : "s"} valid.`;
      note.appendChild(head);

      t.skipped.slice(0, 3).forEach(s => {
        const el = document.createElement("span");
        el.className = "skipped";
        el.textContent = `Row ${s.row} not counted: ${s.reason}.`;
        note.appendChild(el);
      });
      if (t.skipped.length > 3) {
        const more = document.createElement("span");
        more.className = "skipped";
        more.textContent = `+${t.skipped.length - 3} more not counted.`;
        note.appendChild(more);
      }
    }

    if (badge) {
      if (fileSame && receiverRecordExists === true) {
        badge.className = "mode-badge modify";
        badge.textContent = "UPDATE";
      } else if (fileSame && receiverRecordExists === false) {
        badge.className = "mode-badge create";
        badge.textContent = "NEW";
      } else {
        badge.className = "mode-badge hidden";
        badge.textContent = "";
      }
    }
    if (modeNote) {
      if (fileSame && receiverRecordExists === true) {
        modeNote.textContent = "Already in Receiver Records. Sending updates it.";
      } else if (fileSame && receiverRecordExists === false) {
        modeNote.textContent = "Not in Receiver Records yet. Sending adds it.";
      } else {
        modeNote.textContent = "";
      }
    }

    show(totalsBlock);
    hide(totalsEmpty);
    if (fresh) totalsBlock.classList.remove("is-stale");
    else totalsBlock.classList.add("is-stale");
  } else {
    hide(totalsBlock);
    show(totalsEmpty);
    if (badge) {
      badge.className = "mode-badge hidden";
      badge.textContent = "";
    }
  }

  const blocker = getSendBlocker_();
  const blockerEl = $("sendBlocker");
  if (blockerEl) {
    if (blocker && hasFile) {
      blockerEl.textContent = blocker;
      show(blockerEl);
    } else {
      blockerEl.textContent = "";
      hide(blockerEl);
    }
  }

  const canSend = lastAuthOk && !blocker && !busy;
  setButtonEnabled($("create_send"), canSend);
  setStepBadge_($("step3Badge"), 3, false);

  const sendPanel = $("createPanel");
  if (sendPanel) {
    if (canSend) sendPanel.classList.add("is-active");
    else sendPanel.classList.remove("is-active");
  }
}

// Kept for compatibility with older call sites.
function refreshSendButtonsState() {
  refreshUI();
}
