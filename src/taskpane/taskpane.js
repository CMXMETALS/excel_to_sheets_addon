/* global Office, Excel */
/*
  Main entrypoint:
  - wires up DOM events in Office.onReady
  - registers Excel events (tab switch, cell edit, tab rename) — all handled
    locally, none of them call Google Sheets
  - starts the Cloudflare Access sign-in check
*/

let excelEventsRegistered = false;

function debounce_(fn, waitMs) {
  let t = null;
  return function (...args) {
    if (t) clearTimeout(t);
    t = setTimeout(() => fn.apply(this, args), waitMs);
  };
}

// Compare Excel with the last read, shortly after things settle.
const scheduleExcelCheck_ = debounce_(() => {
  checkExcelForChanges_().catch(err => console.warn("excel check error", err));
}, 400);

function excelApiSupported_(version) {
  try {
    return Office.context.requirement.isSetSupported("ExcelApi", version);
  } catch (e) {
    return false;
  }
}

// Each event is registered on its own so an unsupported one cannot block the others.
async function registerOneExcelEvent_(label, attach) {
  try {
    await Excel.run(async (ctx) => {
      attach(ctx.workbook.worksheets);
      await ctx.sync();
    });
    return true;
  } catch (err) {
    console.warn(`Excel event "${label}" not available; Resync and the pre-send check cover it.`, err);
    return false;
  }
}

async function registerExcelTabEvents_() {
  if (excelEventsRegistered) return;
  excelEventsRegistered = true;

  // Tab switch
  await registerOneExcelEvent_("onActivated", (sheets) => {
    sheets.onActivated.add(() => {
      scheduleExcelCheck_();
      return Promise.resolve();
    });
  });

  // Cell edits on any worksheet (only matters for the tab that was read)
  if (excelApiSupported_("1.9")) {
    await registerOneExcelEvent_("onChanged", (sheets) => {
      sheets.onChanged.add((args) => {
        if (lastRead && (!args || !args.worksheetId || args.worksheetId === lastRead.sheetId)) {
          scheduleExcelCheck_();
        }
        return Promise.resolve();
      });
    });
  }

  // Tab rename (newer Excel only)
  if (excelApiSupported_("1.17")) {
    await registerOneExcelEvent_("onNameChanged", (sheets) => {
      sheets.onNameChanged.add(() => {
        scheduleExcelCheck_();
        return Promise.resolve();
      });
    });
  }
}

Office.onReady(() => {
  // auth
  $("cfSignInBtn").addEventListener("click", openCFAccessDialog);

  // sheet reload/select
  $("reloadSheetsBtn").addEventListener("click", loadSheets);
  $("modify_sheet_select").addEventListener("change", onSheetSelected);

  // read + send
  $("create_prefill").addEventListener("click", readFieldsFromExcel_Create);
  $("create_send").addEventListener("click", openSendReviewModal);

  // modal actions
  $("actionModalCancel").addEventListener("click", closeActionModal);
  $("actionModalConfirm").addEventListener("click", confirmPendingSend);

  // Resync: re-check the active tab's name and whether its values changed (Excel only)
  $("refreshActiveSheet").addEventListener("click", async () => {
    await checkExcelForChanges_();
    if (!getSendBlocker_() || !lastRead) setStatus("");
  });

  // receiver details show/hide
  $("detailsToggle").addEventListener("click", () => {
    const btn = $("detailsToggle");
    const open = btn.getAttribute("aria-expanded") === "true";
    btn.setAttribute("aria-expanded", open ? "false" : "true");
    const label = open ? "Show receiver details" : "Hide receiver details";
    btn.setAttribute("aria-label", label);
    btn.setAttribute("title", label);
    refreshUI();
  });

  // Coming back to the pane after working in the grid: re-check locally.
  window.addEventListener("focus", scheduleExcelCheck_);

  hide($("mainContent"));
  setStatus("");
  setAuthView("checking");

  registerExcelTabEvents_()
    .catch(err => console.warn("event registration error", err))
    .then(() => checkAuth({ attempts: 1 }))
    .catch(err => console.warn("startup error", err));
});
