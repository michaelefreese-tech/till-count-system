// validation.gs

// Main validation controller
function ValidateEntry() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("COVER");
  if (!sheet) return false;

  // Reset colors
  sheet.getRange("D10:D19").setBackground(ColorDict["white"]);
  sheet.getRange("G10:G19").setBackground(ColorDict["white"]);
  sheet.getRange("J10:J19").setBackground(ColorDict["white"]);

  // Add input cell borders
  ["D10:D19", "G10:G19", "J10:J19"].forEach(rangeStr => {
    sheet.getRange(rangeStr).setBorder(
      true, true, true, true, true, false,
      "black",
      SpreadsheetApp.BorderStyle.SOLID_THICK
    );
  });

  const till1OK = ValidateTillRange(sheet, 4, 10, "TILL 1");
  const till2OK = ValidateTillRange(sheet, 7, 10, "TILL 2");
  const safeOK = ValidateTillRange(sheet, 10, 10, "SAFE");

  return till1OK && till2OK && safeOK;
}

// Validate one till column block
function ValidateTillRange(sheet, startColumn, startRow, tillName) {
  const labels = [
    "100's",
    "50's",
    "20's",
    "10's",
    "5's",
    "1's",
    "Quarters",
    "Dimes",
    "Nickels",
    "Pennys"
  ];

  for (let i = 0; i < labels.length; i++) {
    const cell = sheet.getRange(startRow + i, startColumn);
    const value = cell.getValue();

    if (value === "" || value === null) {
      SpreadsheetApp.getUi().alert("⛔ Missing value: " + tillName + " : " + labels[i]);
      cell.setBackground("red");
      return false;
    }

    if (isNaN(Number(value))) {
      SpreadsheetApp.getUi().alert("⛔ Invalid number: " + tillName + " : " + labels[i]);
      cell.setBackground("red");
      return false;
    }

    if (Number(value) < 0) {
      SpreadsheetApp.getUi().alert("⛔ Negative values are not allowed: " + tillName + " : " + labels[i]);
      cell.setBackground("red");
      return false;
    }

    cell.setBackground(ColorDict["white"]);
  }

  return true;
}
