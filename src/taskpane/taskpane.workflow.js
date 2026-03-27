/* global Excel */
/*
  Workflow layer:
  - read from Excel for the active worksheet
  - orchestrate create/modify based on detected workspace state
  - workspace name validation
  - button enable / lock-out logic
*/

let pendingSendPlan = null;

// ----------------- READ FROM EXCEL -----------------

async function readFieldsFromExcel_Create() {
  if (!selectedSheetId) {
    setStatus("Select a Google Sheets file first.");
    return;
  }
  if (createInFlight) return;

  createInFlight = true;

  setStatus("Reading Reciever From Excel...");
  try {
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      ws.load("name");

      // Header cells
      const headerRanges = {};
      HEADER_FIELDS.forEach(h => {
        const rng = ws.getRange(h.cell);
        rng.load("values");
        headerRanges[h.key] = rng;
      });

      // Line items block (A18:F...)
      const endRow = ITEM_START_ROW + MAX_LINE_ITEMS - 1;
      const lineRange = ws.getRange(`A${ITEM_START_ROW}:F${endRow}`);
      lineRange.load("values");

      await ctx.sync();

      // Fill header inputs
      HEADER_FIELDS.forEach(h => {
        const el = $(h.cId);
        const rng = headerRanges[h.key];
        const v = rng && rng.values && rng.values[0] ? rng.values[0][0] : "";
        if (el) el.value = v == null ? "" : String(v);
      });

      // Track where these Create values came from
      lastReadCreateSheetName = ws.name || "";

      // Line items
      createLineItems = [];
      const values = lineRange.values || [];
      for (let i = 0; i < values.length; i++) {
        const row = values[i] || [];
        const itemVal  = row[0];
        const comVal   = row[1];
        const grossVal = row[2];
        const tareVal  = row[3];
        const costVal  = row[5]; // F column

        const any =
          (itemVal && String(itemVal).trim().length) ||
          (comVal && String(comVal).trim().length) ||
          (grossVal && String(grossVal).trim().length) ||
          (tareVal && String(tareVal).trim().length) ||
          (costVal && String(costVal).trim().length);

        if (any) {
          createLineItems.push({
            item:  itemVal  == null ? "" : String(itemVal),
            com:   comVal   == null ? "" : String(comVal),
            gross: grossVal == null ? "" : String(grossVal),
            tare:  tareVal  == null ? "" : String(tareVal),
            cost:  costVal  == null ? "" : String(costVal)
          });
        }
      }

        renderLineItems();
    });

    setStatus("Fields Copied From Excel. You Can Edit Them Before Sending.");
  } catch (err) {
    console.error(err);
    setStatus("Failed to read Excel fields: " + (err.message || err));
  } finally {
    createInFlight = false;
    invalidateReceiverRecordStatusCache_();
    await evaluateWorkspaceState();
    updateModeBanner();
    updateResyncWarning();
    refreshSendButtonsState();
  }
}

// ----------------- CREATE / MODIFY WORKSPACE -----------------

function isWorkspaceNameFormatValid(name) {
  if (!name) return false;
  return ID_RE.test(name.trim());
}

function buildReceiverRecordPayloadFromUI() {
  const pfEl  = $("c_pf");
  const drEl  = $("c_dr");
  const poEl  = $("c_po");
  const rcvEl = $("c_rcv");

  const purchaseFrom = pfEl  ? String(pfEl.value || "").trim()  : "";
  const dateReceived = drEl  ? String(drEl.value || "").trim()  : "";
  const poNumber     = poEl  ? String(poEl.value || "").trim()  : "";
  const rrNumber     = rcvEl ? String(rcvEl.value || "").trim() : "";

  if (!rrNumber || !dateReceived) {
    return { payload: null, error: "Receiver # and Date Received are required to log RECEIVER RECORDS." };
  }

  // Net Total = sum of (NET * PRICE) for valid, coded line items.
  // IMPORTANT (split-PO nuance): when TARE is blank, the receiver template often has NET entered
  // directly in the GROSS cell. In that case, treat GROSS as NET.
  function parseNumberOrNull(val) {
    if (val == null) return null;
    // allow "$2.38", "1,234.5", etc.
    const s = String(val).replace(/[^0-9.\-]/g, "").trim();
    if (!s) return null;
    const n = Number(s);
    return Number.isFinite(n) ? n : null;
  }

  let netTotal = 0;
  let sawAny = false;

  for (const li of createLineItems) {
    if (!li) continue;

    const hasAny =
      (li.item && li.item.trim().length) ||
      (li.com && li.com.trim().length) ||
      (li.gross && li.gross.trim().length) ||
      (li.tare && li.tare.trim().length) ||
      (li.cost && li.cost.trim().length);
    if (!hasAny) continue;

    const itemCode = (li.item || "").trim();
    if (!itemCode) continue; // only valid, coded line items

    const gross = parseNumberOrNull(li.gross);
    const tare  = parseNumberOrNull(li.tare);  // optional
    const price = parseNumberOrNull(li.cost);
    if (gross == null || price == null) continue;

    const net = (tare == null) ? gross : (gross - tare);
    if (!Number.isFinite(net) || net <= 0) continue;

    netTotal += (net * price);
    sawAny = true;
  }

  // If we couldn't compute it reliably, send null and let the server write a live formula instead.
  const netTotalOut = sawAny ? Number(netTotal.toFixed(2)) : null;

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
  // Header inputs
  const pfEl       = $("c_pf");
  const drEl       = $("c_dr");
  const carrierEl  = $("c_carrier");
  const poEl       = $("c_po");
  const rcvEl      = $("c_rcv");
  const verEl      = $("c_ver");

  const supplier   = pfEl      ? String(pfEl.value || "").trim()      : "";
  const date       = drEl      ? String(drEl.value || "").trim()      : "";
  const carrier    = carrierEl ? String(carrierEl.value || "").trim() : "";
  const poNumber   = poEl      ? String(poEl.value || "").trim()      : "";
  const rrNumber   = rcvEl     ? String(rcvEl.value || "").trim()     : "";
  const verifiedBy = verEl     ? String(verEl.value || "").trim()     : "";

  // If there is no RR or PO, we can't meaningfully update Master.
  if (!rrNumber || !poNumber) {
    return { payload: null, error: null };
  }

  const lines = [];
  let generatedMissingCode = null;

  function parseNumber(val) {
    if (val == null) return "";
    // allow "$2.38", "1,234.5", etc.
    const s = String(val).replace(/[^0-9.\-]/g, "").trim();
    if (!s) return "";
    const n = Number(s);
    return Number.isFinite(n) ? n : "";
  }

  for (const li of createLineItems) {
    const hasAny =
      (li.item && li.item.trim().length) ||
      (li.com && li.com.trim().length) ||
      (li.gross && li.gross.trim().length) ||
      (li.tare && li.tare.trim().length) ||
      (li.cost && li.cost.trim().length);

    if (!hasAny) continue;

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
      itemCode,                      // for server.js normalizedLines
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
      notes: "",            // header-level notes already live on receiver tab
      carrier,
      lines
    },
    error: null
  };
}



function ensureGeneratedRowsHaveCodes() {
  for (const li of createLineItems) {
    const hasAny =
      (li.item && li.item.trim().length) ||
      (li.com && li.com.trim().length) ||
      (li.gross && li.gross.trim().length) ||
      (li.tare && li.tare.trim().length) ||
      (li.cost && li.cost.trim().length);

    if (!hasAny) continue;

    const parsed = parseCommodityForNotes(li.com);
    const itemCode = (li.item || "").trim();

    if (parsed.isGenerated && !itemCode) {
      return `Generated line "${parsed.material || ""}" requires an item code.`;
    }
  }

  return null;
}

function buildSendPlan() {
  if (!selectedSheetId) {
    return { ok: false, message: "Select a Google Sheets file first." };
  }

  const wsNameInput = $("create_workspace_name");
  const name = (wsNameInput && wsNameInput.value || "").trim();

  if (!name) {
    return { ok: false, message: "Workspace name is empty. Rename the Excel tab so it includes -MM-DD-XXX." };
  }
  if (!isWorkspaceNameFormatValid(name)) {
    return { ok: false, message: "Excel tab name must include -MM-DD-XXX (e.g. NOVA -1-17-025 or NOVA - 1-17-025)." };
  }
  if (workspaceState.type === "title-invalid") {
    return { ok: false, message: "Rename the Excel tab to include -MM-DD-XXX before sending." };
  }
  if (!hasAnyCreateData()) {
    return { ok: false, message: "Nothing to send yet. Fill some header or line items first." };
  }

  const currentName = currentActiveWorksheet;
  if (lastReadCreateSheetName && lastReadCreateSheetName !== currentName) {
    return { ok: false, message: 'Excel sheet changed. Click "Read Data From Excel" again before sending.' };
  }

  const generatedError = ensureGeneratedRowsHaveCodes();
  if (generatedError) {
    return { ok: false, message: generatedError };
  }

  // We no longer create a 1:1 receiver tab copy in Sheets. Sending always:
  // 1) updates Master Receivers, and
  // 2) upserts a row in RECEIVER RECORDS.
  return { ok: true, action: "send", workspaceName: name };
}

function openSendReviewModal() {
  const plan = buildSendPlan();
  if (!plan.ok) {
    setStatus(plan.message || "Unable to send.");
    return;
  }

  pendingSendPlan = plan;
  const modal = $("actionModal");
  const title = $("actionModalTitle");
  const body = $("actionModalBody");

  if (title) title.textContent = "Send receiver to Google Sheets";

  if (body) {
    body.textContent =
      "This will update Master Receivers and log the upload in the RECEIVER RECORDS tab " +
      "(Receiver #, PO#, Purchase From, Date Received, Date Uploaded). " +
      "If this receiver already exists in RECEIVER RECORDS (same RR# + same month/day), Date Uploaded will be refreshed.";
  }

  if (modal) modal.classList.remove("hidden");
}

function closeActionModal() {
  const modal = $("actionModal");
  if (modal) modal.classList.add("hidden");
  pendingSendPlan = null;
}

async function confirmPendingSend() {
  if (!pendingSendPlan || sendInFlight) {
    closeActionModal();
    return;
  }

  const plan = pendingSendPlan;
  closeActionModal();

  sendInFlight = true;
  setButtonEnabled($("create_send"), false);
  setGlobalLoading(true, "Sending receiver to Google Sheets…");
  setStatus("Sending to Google Sheets…");

  try {
    // 1) Determine whether this receiver already exists in RECEIVER RECORDS (RR# + month/day)
    const rrResult = buildReceiverRecordPayloadFromUI();
    if (rrResult.error) throw new Error(rrResult.error);

    let recordExists = false;
    try {
      const qs = new URLSearchParams({
        sheetId: selectedSheetId,
        rrNumber: rrResult.payload.rrNumber,
        poNumber: rrResult.payload.poNumber
      });
      if (rrResult.payload.dateReceived) qs.set("dateReceived", rrResult.payload.dateReceived);
      const stRes = await fetch(`${BACKEND}/api/receiverRecordStatus?${qs.toString()}`, {
        credentials: "include"
      });
      if (stRes.ok) {
        const stJson = await stRes.json();
        recordExists = !!stJson.exists;
      }
    } catch (e) {
      // Non-fatal: we'll still upsert the record
      console.warn("receiverRecordStatus failed:", e);
    }

    // 2) Update Master Receivers (existing logic stays the same)
    const masterResult = buildMasterPayloadFromUI();
    if (masterResult.error) throw new Error(masterResult.error);

    let masterUpdated = false;
    let masterFailStatus = null;

    if (masterResult.payload) {
      try {
        const mrRes = await fetch(`${BACKEND}/api/updateMasterFromReceiver`, {
          method: "POST",
          credentials: "include",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(masterResult.payload)
        });
        masterUpdated = mrRes.ok;
        if (!mrRes.ok) masterFailStatus = mrRes.status;
      } catch (e) {
        masterUpdated = false;
        masterFailStatus = e.message || String(e);
        console.warn("updateMasterFromReceiver error:", e);
      }
    }

    // 3) Upsert RECEIVER RECORDS row (server stamps Date Uploaded)
    const rrUpRes = await fetch(`${BACKEND}/api/upsertReceiverRecord`, {
      method: "POST",
      credentials: "include",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(rrResult.payload)
    });
    if (!rrUpRes.ok) throw new Error("Receiver Records update failed: " + rrUpRes.status);

    const modeText = recordExists ? "updated existing upload record" : "created new upload record";

    // Keep UI state in sync immediately (and invalidate any cached status checks)
    try {
      workspaceState.receiverRecordExists = true;
      if (typeof invalidateReceiverRecordStatusCache_ === "function") invalidateReceiverRecordStatusCache_();
      if (typeof updateModeBanner === "function") updateModeBanner();
      if (typeof refreshSendButtonsState === "function") refreshSendButtonsState();
    } catch (_) {}

    if (!masterResult.payload) {
      setStatus(`Done: ${modeText}. (Master not updated: missing RR#/PO#/lines.)`);
    } else if (!masterUpdated) {
      setStatus(`Done: ${modeText}, but Master Receivers update failed (${masterFailStatus}).`);
    } else {
      setStatus(`Done: ${modeText} and updated Master Receivers.`);
    }

    await new Promise(r => setTimeout(r, 250));
    await onSheetSelected();
  } catch (err) {
    console.error(err);
    setStatus("Send failed: " + (err.message || err));
  } finally {
    sendInFlight = false;
    setGlobalLoading(false);
    refreshSendButtonsState();
  }
}



function refreshSendButtonsState() {
  let sendEnabled = false;

  // Pre-conditions: auth + sheet selected
  if (!lastAuthOk) {
    setButtonEnabled($("create_send"), false);
    setButtonEnabled($("create_prefill"), false);
    setSectionLocked($("createPanel"), true);
    return;
  }

  setSectionLocked($("createPanel"), !selectedSheetId);

  if (!selectedSheetId) {
    setButtonEnabled($("create_send"), false);
    setButtonEnabled($("create_prefill"), false);
    return;
  }

  const wsName = ($("create_workspace_name") && $("create_workspace_name").value || "").trim();
  const nameValid = isWorkspaceNameFormatValid(wsName);
  const currentName = currentActiveWorksheet;
  const needsResync =
    hasAnyCreateData() &&
    lastReadCreateSheetName &&
    currentName &&
    lastReadCreateSheetName !== currentName;

  const generatedError = ensureGeneratedRowsHaveCodes();

  // allow prefill regardless of format, but disable while read/send in flight
  setButtonEnabled($("create_prefill"), !!selectedSheetId && !createInFlight && !sendInFlight);

  if (!sendInFlight &&
      selectedSheetId &&
      // Title can be invalid, but we can still send once Receiver # + PO# are present.
      (workspaceState.type !== "title-invalid" || ((($("c_rcv") && $("c_rcv").value) || "").trim() && (($("c_po") && $("c_po").value) || "").trim())) &&
      nameValid &&
      hasAnyCreateData() &&
      !needsResync &&
      !generatedError) {
    if (workspaceState.type === "existing" || workspaceState.type === "new-valid") {
      sendEnabled = true;
    }
  }

  setButtonEnabled($("create_send"), sendEnabled);
}
