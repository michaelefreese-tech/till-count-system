// editTrigger.gs

// Main onEdit trigger — intercepts user changes

function onEdit(e) {

  const ss = e.source;
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();
  const range = e.range;
  const editedCell = range.getA1Notation();
  const row = range.getRow();
  const col = range.getColumn();

  const props = PropertiesService.getDocumentProperties();
  const isLocked = props.getProperty("docLocked") === "true";

  // If document is not locked, allow edit
  if (!isLocked) return;

  // ================================
  // SHEET: SPLASH
  // ================================

  if (sheetName === "SPLASH") {
    if (editedCell !== "K13") {
      SpreadsheetApp.getUi().alert("⛔ You don't have permission to edit this cell: " + editedCell);
      Restore();
    }
    return;
  }

  // ================================
  // SHEET: COVER
  // ================================

  if (sheetName === "COVER") {
    const allowedCells = Info["COVER"].ALLOWEDCELLS || [];

    if (!allowedCells.includes(editedCell)) {
      SpreadsheetApp.getUi().alert("⛔ You don't have permission to edit this cell: " + editedCell);
      Restore();
      UpdateCoverSheet();
      return;
    }

    // Search field is allowed but does not live-sync to a shift sheet
    if (editedCell === "J24") {
      return;
    }

    // Only denomination entry cells should live-sync to the selected shift sheet
    const isDataEntryCell =
      (col === 4 || col === 7 || col === 10) &&
      row >= 10 &&
      row <= 19;

    if (!isDataEntryCell) {
      return;
    }

    const shiftType = sheet
      .getRange(CurrentShiftIndicator)
      .getValue()
      .toString()
      .trim()
      .toUpperCase();

    // SEARCH mode is view-only
    if (shiftType === "SEARCH") {
      SpreadsheetApp.getUi().alert(
        "⛔ Search mode is view-only.\nClear search results and return to a live shift to edit denominations."
      );
      UpdateCoverSheet();
      return;
    }

    const targetSheet = ss.getSheetByName(shiftType);
    if (!targetSheet) {
      Logger.log("Target shift sheet not found: " + shiftType);
      return;
    }

    // Live sync cover entry to active shift sheet
    targetSheet.getRange(row, col).setValue(range.getValue());

    // Refresh active cover view
    UpdateCoverSheet();
    return;
  }

  // ================================
  // SHEET: INFO
  // ================================

  if (sheetName === "INFO") {
    const allowedCells = Info["INFO"].ALLOWEDCELLS || [];

    if (!allowedCells.includes(editedCell)) {
      SpreadsheetApp.getUi().alert("⛔ You don't have permission to edit this INFO cell: " + editedCell);
      Restore();
      GetWeekInfo();
      UpdateCoverSheet();
      return;
    }

    // If period number or first day changes, rebuild week info
    if (editedCell === "D2" || editedCell === "D3") {
      GetWeekInfo();
      Restore();
      UpdateCoverSheet();
    }

    return;
  }

  // ================================
  // SHEETS: OPEN / SHIFT / CLOSE / SEARCH
  // ================================

  if (["OPEN", "SHIFT", "CLOSE", "SEARCH"].includes(sheetName)) {
    SpreadsheetApp.getUi().alert("⛔ This sheet is locked.");

    // Rebuild static structure, then redraw COVER from whichever mode is active
    Restore();
    UpdateCoverSheet();
    return;
  }

  // ================================
  // SHEETS: WEEK_1 / WEEK_2 / WEEK_3 / WEEK_4
  // ================================

  if (["WEEK_1", "WEEK_2", "WEEK_3", "WEEK_4"].includes(sheetName)) {
    SpreadsheetApp.getUi().alert("⛔ This sheet is locked.");

    Restore();
    return;
  }

  // ================================
  // ALL OTHER SHEETS
  // ================================

  SpreadsheetApp.getUi().alert("⛔ This sheet is locked.");
  Restore();
}

function onChange(e) {

  const props = PropertiesService.getDocumentProperties();
  const isLocked = props.getProperty("docLocked") === "true";
  if (!isLocked) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existingNames = ss.getSheets().map(s => s.getName());

  REQUIRED_SHEETS.forEach(name => {
    if (!existingNames.includes(name)) {
      ss.insertSheet(name);
    }
  });

  // If anything structural changed OR sheets were missing → full rebuild
  Restore();
  GetWeekInfo();
  UpdateCoverSheet();
  ApplySheetProtections();
}

function InstallChangeTrigger() {

  const ss = SpreadsheetApp.getActive();

  const existing = ScriptApp.getProjectTriggers();
  existing.forEach(t => {
    if (t.getHandlerFunction() === "onChange") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("onChange")
    .forSpreadsheet(ss)
    .onChange()
    .create();
}

function EmergencyUnlock() {

  const props = PropertiesService.getDocumentProperties();
  props.setProperty("docLocked", "false");
}
