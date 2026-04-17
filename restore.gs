// restore.gs

// Full sheet structure restoration using Info dictionary
function Restore() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const sheetInfo = Info[sheetName];
    if (!sheetInfo) return;

    // Merge ranges
    if (sheetInfo.MERGE) {
      sheetInfo.MERGE.forEach(rangeStr => {
        const range = sheet.getRange(rangeStr);
        try {
          if (range.isPartOfMerge()) range.breakApart();
          range.merge();
        } catch (err) {
          Logger.log(`Could not merge ${sheetName}!${rangeStr}: ${err}`);
        }
      });
    }

    if (sheetInfo.MERGERIGHT) {
      sheetInfo.MERGERIGHT.forEach(rangeStr => {
        const range = sheet.getRange(rangeStr);
        try {
          if (range.isPartOfMerge()) range.breakApart();
          range.merge();
        } catch (err) {
          Logger.log(`Could not merge ${sheetName}!${rangeStr}: ${err}`);
        }
      });
    }

    if (sheetInfo.MERGELEFT) {
      sheetInfo.MERGELEFT.forEach(rangeStr => {
        const range = sheet.getRange(rangeStr);
        try {
          if (range.isPartOfMerge()) range.breakApart();
          range.merge();
        } catch (err) {
          Logger.log(`Could not merge ${sheetName}!${rangeStr}: ${err}`);
        }
      });
    }

    if (sheetInfo.MERGECENTER) {
      sheetInfo.MERGECENTER.forEach(rangeStr => {
        const range = sheet.getRange(rangeStr);
        try {
          if (range.isPartOfMerge()) range.breakApart();
          range.merge();
        } catch (err) {
          Logger.log(`Could not merge ${sheetName}!${rangeStr}: ${err}`);
        }
      });
    }

    // Apply alignment
    if (sheetInfo.MERGE) {
      sheetInfo.MERGE.forEach(rangeStr => {
        sheet.getRange(rangeStr)
          .setHorizontalAlignment("left")
          .setVerticalAlignment("middle");
      });
    }

    if (sheetInfo.RIGHT) {
      sheetInfo.RIGHT.forEach(rangeStr => {
        sheet.getRange(rangeStr)
          .setHorizontalAlignment("right")
          .setVerticalAlignment("middle");
      });
    }

    if (sheetInfo.MERGERIGHT) {
      sheetInfo.MERGERIGHT.forEach(rangeStr => {
        sheet.getRange(rangeStr)
          .setHorizontalAlignment("right")
          .setVerticalAlignment("middle");
      });
    }

    if (sheetInfo.MERGELEFT) {
      sheetInfo.MERGELEFT.forEach(rangeStr => {
        sheet.getRange(rangeStr)
          .setHorizontalAlignment("left")
          .setVerticalAlignment("middle");
      });
    }

    if (sheetInfo.MERGECENTER) {
      sheetInfo.MERGECENTER.forEach(rangeStr => {
        sheet.getRange(rangeStr)
          .setHorizontalAlignment("center")
          .setVerticalAlignment("middle");
      });
    }

    // Apply background colors
    if (sheetInfo.COLORS) {
      for (const colorKey in sheetInfo.COLORS) {
        const hex = ColorDict[colorKey.toLowerCase()];
        if (!hex) continue;

        sheetInfo.COLORS[colorKey].forEach(rangeStr => {
          sheet.getRange(rangeStr).setBackground(hex);
        });
      }
    }

    // Apply borders
    if (sheetInfo.OUTLINETHIN) {
      sheetInfo.OUTLINETHIN.forEach(rangeStr => {
        sheet.getRange(rangeStr).setBorder(
          true, true, true, true, false, false,
          "black",
          SpreadsheetApp.BorderStyle.SOLID
        );
      });
    }

    if (sheetInfo.OUTLINEMED) {
      sheetInfo.OUTLINEMED.forEach(rangeStr => {
        sheet.getRange(rangeStr).setBorder(
          true, true, true, true, false, false,
          "black",
          SpreadsheetApp.BorderStyle.SOLID_MEDIUM
        );
      });
    }

    // Apply strings
    if (sheetInfo.STRING) {
      for (const cell in sheetInfo.STRING) {
        sheet.getRange(cell)
          .setValue(sheetInfo.STRING[cell])
          .setFontFamily("Arial")
          .setFontSize(10);
      }
    }

    // Apply formulas
    if (sheetInfo.FORMULAS) {
      for (const cell in sheetInfo.FORMULAS) {
        sheet.getRange(cell).setFormula(sheetInfo.FORMULAS[cell]);
      }
    }

    // Number formatting
    if (sheetInfo.FORMAT) {
      for (const formatType in sheetInfo.FORMAT) {
        sheetInfo.FORMAT[formatType].forEach(rangeStr => {
          const range = sheet.getRange(rangeStr);

          if (formatType === "currency") {
            range.setNumberFormat("$#,##0.00");
          }

          if (formatType === "currency_round") {
            range.setNumberFormat("$#,##0");
          }

          if (formatType === "number") {
            range.setNumberFormat("0.00");
          }

          if (formatType === "date") {
            range.setNumberFormat("MM/dd/yyyy");
          }

          if (formatType === "datetime") {
            range.setNumberFormat("MM/dd/yyyy hh:mm AM/PM");
          }
        });
      }
    }

    // Sheet-specific formatting
    if (sheetName === "SPLASH") {
      sheet.getRange("C4:O6").setFontSize(26).setFontWeight("bold");
      sheet.getRange("C8:O9").setFontSize(18).setFontWeight("bold");
      sheet.getRange("G13:J14").setFontSize(11).setFontWeight("bold");
      sheet.getRange("K13:M14").setFontSize(12);
    }

    if (sheetName === "COVER") {
      // Till 1, Till 2, Safe
      sheet.getRange("C8").setFontWeight("bold");
      sheet.getRange("F8").setFontWeight("bold");
      sheet.getRange("I8").setFontWeight("bold");

      // Shift info box
      sheet.getRange("N13").setFontSize(8);
      sheet.getRange("N15").setFontSize(8);
      sheet.getRange("N17").setFontSize(8);
      sheet.getRange("N19").setFontSize(8);
      sheet.getRange("N21").setFontSize(8);

      // Employee id, Date to search
      sheet.getRange("H24").setFontWeight("bold");
      sheet.getRange("B24").setFontWeight("bold");
    }

    if (sheetName === "INFO") {
      sheet.getRange("B2").setFontWeight("bold");
      sheet.getRange("B3").setFontWeight("bold");
      sheet.getRange("B5").setFontWeight("bold");
    }

    if (sheetName === "WEEK_1" || sheetName === "WEEK_2" || sheetName === "WEEK_3" || sheetName === "WEEK_4") {
      sheet.getRange("C3:AF4").setFontWeight("bold").setFontSize(14);
      sheet.getRange("C16:AF17").setFontWeight("bold").setFontSize(14);
      sheet.getRange("C29:AF30").setFontWeight("bold").setFontSize(14);

      sheet.getRange("C5:AF5").setFontWeight("bold");
      sheet.getRange("C18:AF18").setFontWeight("bold");
      sheet.getRange("C31:AF31").setFontWeight("bold");

      sheet.getRange("A7:A13").setFontWeight("bold");
      sheet.getRange("A20:A26").setFontWeight("bold");
      sheet.getRange("A33:A39").setFontWeight("bold");

      sheet.getRange("B6").setFontWeight("bold");
      sheet.getRange("AG6:AI6").setFontWeight("bold");
      sheet.getRange("AG19:AI19").setFontWeight("bold");
      sheet.getRange("AG32:AI32").setFontWeight("bold");
    }
  });
}
