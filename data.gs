// data.gs
//----------------------------------------------------------------------------------------------------
// FUNCTIONS FOR TRANSFERING INFO FROM SHIFTS SHEETS TO WEEK SHEETS
//----------------------------------------------------------------------------------------------------

// Transfers data for all shifts (OPEN, SHIFT, CLOSE) from their respective sheets to the current WEEK sheet

function TransferData() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const coverSheet = ss.getSheetByName("COVER");
  const currentMode = GetCurrentViewMode();
  const validShifts = ["OPEN", "SHIFT", "CLOSE"];
  const selectedWeek = GetCurrentWeek();
  const targetSheet = ss.getSheetByName(selectedWeek);
  const baseColumn = 3; // Column C
  const saveDate = new Date();
  const employeeId = coverSheet ? coverSheet.getRange("D24").getValue() : "";
  const savedShiftNames = [];
  const overwriteConflicts = [];

  if (currentMode === "SEARCH") {
    ui.alert(
      "Save is disabled while viewing search results.\nPlease clear search mode and return to a live shift."
    );
    return;
  }

  if (!selectedWeek || !targetSheet) {
    Logger.log("Could not determine the current week or find the week sheet.");
    return;
  }

  if (!ValidateEntry()) {
    return;
  }

  // First pass: detect any existing saved data that would be overwritten
  for (const shiftType of validShifts) {
    const shiftSheet = ss.getSheetByName(shiftType);
    const targetRow = GetShiftRow(targetSheet, shiftType, saveDate);

    if (!shiftSheet) {
      Logger.log("Sheet for " + shiftType + " not found.");
      continue;
    }

    if (!targetRow) {
      Logger.log("Could not determine the target row for " + shiftType);
      continue;
    }

    const existingPackedData = targetSheet.getRange(targetRow, baseColumn, 1, 30).getValues()[0];
    const existingTotal = targetSheet.getRange(targetRow, 33).getValue(); // AG
    const submittedOn = targetSheet.getRange(targetRow, 34).getValue();   // AH
    const submittedBy = targetSheet.getRange(targetRow, 35).getValue();   // AI

    const alreadyHasPackedData = existingPackedData.some(value => value !== "" && value !== null);
    const alreadyHasMetadata =
      (existingTotal !== "" && existingTotal !== null) ||
      (submittedOn !== "" && submittedOn !== null) ||
      (submittedBy !== "" && submittedBy !== null);

    if (alreadyHasPackedData || alreadyHasMetadata) {
      const submittedOnText = submittedOn instanceof Date && !isNaN(submittedOn)
        ? Utilities.formatDate(submittedOn, Session.getScriptTimeZone(), "MM/dd/yyyy hh:mm a")
        : (submittedOn || "N/A");

      overwriteConflicts.push(
        shiftType +
        " | Existing Total: " + (existingTotal || 0) +
        " | Submitted On: " + submittedOnText +
        " | Submitted By: " + (submittedBy || "N/A")
      );
    }
  }

  // Ask once if any existing submissions would be overwritten
  if (overwriteConflicts.length > 0) {
    const response = ui.alert(
      "Overwrite Existing Submission(s)?",
      "One or more shifts already have saved data for " +
        Utilities.formatDate(new Date(saveDate), Session.getScriptTimeZone(), "MM/dd/yyyy") +
        ".\n\n" +
        overwriteConflicts.join("\n") +
        "\n\nDo you want to overwrite these submission(s)?",
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      Logger.log("Overwrite cancelled for " + selectedWeek);
      return;
    }
  }

  // Second pass: save all shifts
  for (const shiftType of validShifts) {
    const shiftSheet = ss.getSheetByName(shiftType);
    const targetRow = GetShiftRow(targetSheet, shiftType, saveDate);

    if (!shiftSheet) {
      Logger.log("Sheet for " + shiftType + " not found.");
      continue;
    }

    if (!targetRow) {
      Logger.log("Could not determine the target row for " + shiftType);
      continue;
    }

    const till1Range = shiftSheet.getRange("D10:D19").getValues().map(r => r[0]);
    const till2Range = shiftSheet.getRange("G10:G19").getValues().map(r => r[0]);
    const safeRange = shiftSheet.getRange("J10:J19").getValues().map(r => r[0]);
    const packedData = till1Range.concat(till2Range, safeRange);

    for (let i = 0; i < packedData.length; i++) {
      targetSheet.getRange(targetRow, baseColumn + i).setValue(packedData[i]);
    }

    // Submitted on / submitted by
    targetSheet.getRange("AH" + targetRow).setValue(new Date());
    targetSheet.getRange("AI" + targetRow).setValue(employeeId);

    savedShiftNames.push(shiftType);
    Logger.log("Transferred " + shiftType + " data to " + selectedWeek + " row " + targetRow);
  }

  // Clear only shifts that were actually saved
  ClearShiftSheets(savedShiftNames);
  UpdateCoverSheet();
}

// Determines the correct row for the shift in the WEEK sheet

function GetShiftRow(sheet, shiftType, searchDate) {

  const shiftInfo = { "OPEN": 7, "SHIFT": 20, "CLOSE": 33 }[shiftType];
  if (!shiftInfo) return null;

  const dateColumn = sheet.getRange("B" + shiftInfo + ":B" + (shiftInfo + 6)).getValues().flat();
  const formattedSearchDate = Utilities.formatDate(new Date(searchDate), Session.getScriptTimeZone(), "MM/dd/yyyy");

  for (let i = 0; i < dateColumn.length; i++) {
    const cellValue = dateColumn[i];
    if (!cellValue) continue;

    const formattedCellDate = Utilities.formatDate(new Date(cellValue), Session.getScriptTimeZone(), "MM/dd/yyyy");
    Logger.log(
      shiftType +
      " checking row " + (shiftInfo + i) +
      " | sheet date: " + formattedCellDate +
      " | search date: " + formattedSearchDate
    );

    if (formattedCellDate === formattedSearchDate) {
      return shiftInfo + i;
    }
  }

  Logger.log("No matching date found for " + shiftType + " on " + formattedSearchDate);
  return null;
}

// Returns the week sheet name for a given date

function GetCurrentWeek(searchDate = null) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const infoSheet = ss.getSheetByName("INFO");
  const startValue = infoSheet.getRange("D3").getValue();
  const baseDate = new Date(startValue);
  const targetDate = searchDate ? new Date(searchDate) : new Date();

  if (!(baseDate instanceof Date) || isNaN(baseDate)) return null;
  if (!(targetDate instanceof Date) || isNaN(targetDate)) return null;

  const normalizedTarget = new Date(targetDate);
  normalizedTarget.setHours(0, 0, 0, 0);

  for (let i = 0; i < 4; i++) {
    const weekStart = new Date(baseDate);
    weekStart.setDate(baseDate.getDate() + (i * 7));
    weekStart.setHours(0, 0, 0, 0);

    const weekEnd = new Date(weekStart);
    weekEnd.setDate(weekStart.getDate() + 6);
    weekEnd.setHours(23, 59, 59, 999);

    if (normalizedTarget >= weekStart && normalizedTarget <= weekEnd) {
      return "WEEK_" + (i + 1);
    }
  }

  return null;
}

//----------------------------------------------------------------------------------------------------
// FUNCTIONS FOR SEARCHING THE DATABASE AND RECALLING INFO TO THE COVER PAGE
//----------------------------------------------------------------------------------------------------

// Retrieves data from WEEK sheet based on the date in J24 and populates the SEARCH sheet

function SearchData() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const coverSheet = ss.getSheetByName("COVER");
  const searchSheet = ss.getSheetByName("SEARCH");
  const searchDate = coverSheet.getRange("J24").getValue();
  const selectedWeek = GetCurrentWeek(searchDate);
  const sourceSheet = ss.getSheetByName(selectedWeek);
  const validShifts = ["OPEN", "SHIFT", "CLOSE"];
  const currentMode = GetCurrentViewMode();

  if (!searchDate) {
    SpreadsheetApp.getUi().alert("Please enter a date to search.");
    return;
  }

  if (!searchSheet) {
    SpreadsheetApp.getUi().alert("SEARCH sheet not found.");
    Logger.log("SEARCH sheet not found.");
    return;
  }

  if (!selectedWeek || !sourceSheet) {
    SpreadsheetApp.getUi().alert("Could not determine the week sheet for that date.");
    Logger.log("Could not determine the week or find the week sheet.");
    return;
  }

  // Remember the last live shift before entering SEARCH mode
  if (VALID_LIVE_SHIFTS.includes(currentMode)) {
    SetLastLiveShift(currentMode);
  }

  // Clear previous SEARCH results first
  Info["SEARCH"].DATACELLS.forEach(rangeStr => {
    searchSheet.getRange(rangeStr).clearContent();
  });

  let restoredCount = 0;

  validShifts.forEach(shiftType => {
    const row = GetShiftRow(sourceSheet, shiftType, searchDate);

    Logger.log("SEARCHING " + shiftType + " | WEEK: " + selectedWeek + " | ROW: " + row);

    if (!row) {
      Logger.log("No row found for " + shiftType + " on " + searchDate);
      return;
    }

    const rowData = sourceSheet.getRange(row, 3, 1, 30).getValues()[0];
    const till1 = rowData.slice(0, 10);
    const till2 = rowData.slice(10, 20);
    const safe = rowData.slice(20, 30);

    // SEARCH sheet uses same layout as OPEN / SHIFT / CLOSE
    searchSheet.getRange("D10:D19").setValues(till1.map(v => [v]));
    searchSheet.getRange("G10:G19").setValues(till2.map(v => [v]));
    searchSheet.getRange("J10:J19").setValues(safe.map(v => [v]));

    restoredCount++;
  });

  coverSheet.getRange(CurrentShiftIndicator).setValue("SEARCH");
  UpdateCoverSheet();

  SpreadsheetApp.getUi().alert(
    "Search complete.\nWeek: " + selectedWeek + "\nHistorical shifts loaded into SEARCH: " + restoredCount
  );
}

// Clears OPEN, SHIFT, and CLOSE sheet entry cells

function ClearShiftSheets(shiftNames = ["OPEN", "SHIFT", "CLOSE"]) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  shiftNames.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    const dataRanges = Info[sheetName]?.DATACELLS || [];

    if (!sheet) return;

    dataRanges.forEach(rangeStr => {
      sheet.getRange(rangeStr).clearContent();
    });
  });
}

// Clear search data and exit SEARCH mode

function ClearSearchData() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const coverSheet = ss.getSheetByName("COVER");
  const searchSheet = ss.getSheetByName("SEARCH");

  if (!searchSheet || !coverSheet) return;

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Clear Search Results",
    "This will clear the current search results and return to your last live shift.\n\nContinue?",
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  // Clear only SEARCH sheet data
  const dataRanges = Info["SEARCH"]?.DATACELLS || [];
  dataRanges.forEach(rangeStr => {
    searchSheet.getRange(rangeStr).clearContent();
  });

  // Restore last live shift (OPEN / SHIFT / CLOSE)
  const lastLiveShift = GetLastLiveShift();
  coverSheet.getRange(CurrentShiftIndicator).setValue(lastLiveShift);

  UpdateCoverSheet();

  ui.alert("Search cleared. Returned to " + lastLiveShift + " mode.");
}

//----------------------------------------------------------------------------------------------------
// FUNCTIONS FOR CREATING RANDOM INFO FOR ALL SHIFTS FOR DEBUGGING AND TESTING PURPOSES
//----------------------------------------------------------------------------------------------------

// Creates random data in the defined data cells for each shift sheet
function CreateData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shiftNames = ["OPEN", "SHIFT", "CLOSE"];

  shiftNames.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    const dataRanges = Info[sheetName]?.DATACELLS || [];

    if (!sheet || dataRanges.length === 0) return;

    dataRanges.forEach(rangeStr => {
      const range = sheet.getRange(rangeStr);
      const numRows = range.getNumRows();
      const values = [];
      const randomValue1 = parseFloat((Math.random() * (250 - 0.01) + 0.01).toFixed(2));

      for (let i = 0; i < numRows; i++) {
        const randomValue = parseFloat((Math.random() * (randomValue1 - 0.01) + 0.01).toFixed(2));
        values.push([randomValue]);
      }

      range.setValues(values);
    });
  });

  UpdateCoverSheet();
  Logger.log("Random data created in OPEN, SHIFT, and CLOSE sheets.");
}

// Creates test data for yesterday in the proper WEEK sheet only if it does not already exist
function CreateYesterdayWeekData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const coverSheet = ss.getSheetByName("COVER");
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  yesterday.setHours(0, 0, 0, 0);

  const selectedWeek = GetCurrentWeek(yesterday);
  const targetSheet = ss.getSheetByName(selectedWeek);
  const validShifts = ["OPEN", "SHIFT", "CLOSE"];
  const baseColumn = 3; // Column C
  const employeeId = coverSheet ? coverSheet.getRange("D24").getValue() : "TEST";

  if (!selectedWeek || !targetSheet) {
    SpreadsheetApp.getUi().alert("Could not determine the correct week sheet for yesterday.");
    return;
  }

  let results = [];

  validShifts.forEach(shiftType => {
    const targetRow = GetShiftRow(targetSheet, shiftType, yesterday);

    if (!targetRow) {
      results.push(`${shiftType}: no matching row found`);
      return;
    }

    const existingData = targetSheet.getRange(targetRow, baseColumn, 1, 30).getValues()[0];
    const hasData = existingData.some(value => value !== "" && value !== null);

    if (hasData) {
      results.push(`${shiftType}: already had data`);
      return;
    }

    const packedData = [];
    const randomMax = parseFloat((Math.random() * (250 - 25) + 25).toFixed(2));

    for (let i = 0; i < 30; i++) {
      packedData.push(parseFloat((Math.random() * randomMax).toFixed(2)));
    }

    for (let i = 0; i < packedData.length; i++) {
      targetSheet.getRange(targetRow, baseColumn + i).setValue(packedData[i]);
    }

    targetSheet.getRange("AH" + targetRow).setValue(new Date());
    targetSheet.getRange("AI" + targetRow).setValue(employeeId || "TEST");

    results.push(`${shiftType}: created test data`);
  });

  SpreadsheetApp.getUi().alert(
    "Yesterday test-data check complete for " +
    Utilities.formatDate(yesterday, Session.getScriptTimeZone(), "MM/dd/yyyy") +
    " in " + selectedWeek + ".\n\n" +
    results.join("\n")
  );
}
//----------------------------------------------------------------------------------------------------
// FUNCTIONS FOR SETTING SHIFT INDICATOR
//----------------------------------------------------------------------------------------------------

// Handles button clicks and changes shift indicator

function ChangeShiftIndicator(shiftType) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const coverSheet = ss.getSheetByName("COVER");

  if (!coverSheet) return;

  const normalized = (shiftType || "").toUpperCase().trim();

  // If switching to a live shift, remember it
  if (VALID_LIVE_SHIFTS.includes(normalized)) {
    SetLastLiveShift(normalized);
  }

  coverSheet.getRange(CurrentShiftIndicator).setValue(normalized);
  UpdateCoverSheet();
}

function SetOpenShift() {
  ChangeShiftIndicator("OPEN");
}

function SetShift() {
  ChangeShiftIndicator("SHIFT");
}

function SetCloseShift() {
  ChangeShiftIndicator("CLOSE");
}
