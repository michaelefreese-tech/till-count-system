// cover.gs

// Update COVER sheet visuals and values

function UpdateCoverSheet() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const coverSheet = ss.getSheetByName("COVER");

  if (!coverSheet) {
    Logger.log("COVER sheet not found.");
    return;
  }

  const shiftTypeLower = coverSheet
    .getRange(CurrentShiftIndicator)
    .getValue()
    .toString()
    .trim()
    .toLowerCase();

  const shiftTypeUpper = shiftTypeLower.toUpperCase();

  // Apply menu colors based on current shift / mode
  const menuRanges = Info["COVER"]?.MENU?.[shiftTypeLower] || [];
  menuRanges.forEach(rangeStr => {
    const color = ColorDict[shiftTypeLower];
    if (color) {
      coverSheet.getRange(rangeStr).setBackground(color);
    }
  });

  // Apply configured COVER colors
  const coverColors = Info["COVER"]?.COLORS || {};
  for (const colorKey in coverColors) {
    const hex = ColorDict[colorKey.toLowerCase()];
    if (!hex) continue;

    coverColors[colorKey].forEach(rangeStr => {
      coverSheet.getRange(rangeStr).setBackground(hex);
    });
  }

  // Pull current mode data onto COVER
  const sourceSheet = ss.getSheetByName(shiftTypeUpper);
  if (!sourceSheet) {
    Logger.log("Source sheet " + shiftTypeUpper + " not found.");
    return;
  }

  const dataRanges = Info[shiftTypeUpper]?.DATACELLS || [];
  dataRanges.forEach(rangeStr => {
    const sourceRange = sourceSheet.getRange(rangeStr);
    const targetRange = coverSheet.getRange(rangeStr);
    targetRange.setValues(sourceRange.getValues());
  });

  // SEARCH mode is view-only and should not show previous-shift date logic
  if (shiftTypeUpper === "SEARCH") {
    coverSheet.getRange("O13").clearContent();
    return;
  }

  let previousShiftDate = "";
  const currentDateValue = coverSheet.getRange("O17").getValue();

  // SHIFT and CLOSE use same-day previous-shift date
  if (shiftTypeUpper === "SHIFT" || shiftTypeUpper === "CLOSE") {
    previousShiftDate = currentDateValue || "";
  }

  // OPEN uses previous day's CLOSE date from the proper WEEK sheet row
  if (shiftTypeUpper === "OPEN") {
    if (currentDateValue instanceof Date && !isNaN(currentDateValue)) {
      const searchDate = new Date(currentDateValue);
      searchDate.setDate(searchDate.getDate() - 1);
      searchDate.setHours(0, 0, 0, 0);

      previousShiftDate = searchDate;

      const previousWeekName = GetCurrentWeek(searchDate);
      const previousWeekSheet = ss.getSheetByName(previousWeekName);

      if (previousWeekSheet) {
        const previousRow = GetShiftRow(previousWeekSheet, "CLOSE", searchDate);

        if (previousRow) {
          previousShiftDate = previousWeekSheet.getRange("B" + previousRow).getValue() || searchDate;
        } else {
          Logger.log("No previous CLOSE row found for OPEN on " + searchDate);
        }
      } else {
        Logger.log("No previous week sheet found for OPEN on " + searchDate);
      }
    } else {
      previousShiftDate = "";
    }
  }

  // Update side panel values that are not already formula-driven
  coverSheet.getRange("O13").setValue(previousShiftDate || "");
}
