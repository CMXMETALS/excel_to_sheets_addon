/* global Office, Excel */
/*
  Main entrypoint:
  - wires up DOM events in Office.onReady
  - starts active worksheet polling with blocking overlay on tab changes
  - routes send button to review modal handled in workflow.js
*/

let excelEventsRegistered = false;

function debounce_(fn, waitMs) {
  let t = null;
  return function (...args) {
    if (t) clearTimeout(t);
    t = setTimeout(() => fn.apply(this, args), waitMs);
  };
}

async function registerExcelTabEvents_() {
  if (excelEventsRegistered) return;
  try {
    await Excel.run(async (ctx) => {
      // Fires when user changes the active worksheet (tab switch)
      ctx.workbook.worksheets.onActivated.add(() => {
        detectActiveWorksheet(true).catch(err => console.warn("tab activated detect error", err));
      });
      await ctx.sync();
    });
    excelEventsRegistered = true;
  } catch (err) {
    console.warn("Worksheet activation event not available; rely on manual refresh + key events", err);
  }
}

Office.onReady(() => {
  // auth
  $("cfSignInBtn").addEventListener("click", openCFAccessDialog);


  // sheet reload/select
  $("reloadSheetsBtn").addEventListener("click", loadSheets);
  $("modify_sheet_select").addEventListener("change", onSheetSelected);

  // create panel
  $("create_prefill").addEventListener("click", readFieldsFromExcel_Create);
  $("create_send").addEventListener("click", openSendReviewModal);
  $("create_add_line").addEventListener("click", () => addLineItem());
  $("create_workspace_name").addEventListener("input", refreshSendButtonsState);

  // Key fields affect Create vs Modify; re-evaluate on change (no constant polling)
  const onKeyChange = debounce_(async () => {
    invalidateReceiverRecordStatusCache_();
    await evaluateWorkspaceState();
    updateModeBanner();
    updateResyncWarning();
    refreshSendButtonsState();
  }, 250);
  ["c_rcv", "c_po", "c_dr"].forEach(id => {
    const el = $(id);
    if (el) el.addEventListener("input", onKeyChange);
  });

  // modal actions
  $("actionModalCancel").addEventListener("click", closeActionModal);
  $("actionModalConfirm").addEventListener("click", confirmPendingSend);

  // Optional manual refresh button – still works as a quick resync
  const refreshBtn = $("refreshActiveSheet");
  if (refreshBtn) {
    refreshBtn.addEventListener("click", async () => {
      await detectActiveWorksheet(true);
      refreshSendButtonsState();
    });
  }

  hide($("mainContent"));
  setStatus("Connecting...");
  registerExcelTabEvents_()
    .then(() => checkAuth())
    .catch(err => console.warn("startup error", err));

  // initial detect + register tab-change handler (no constant polling)
  registerExcelTabEvents_()
    .then(() => detectActiveWorksheet(true))
    .catch(err => console.warn("initial detect error", err));
});
