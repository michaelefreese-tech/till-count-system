// weeks.gs

// Generate 4-week calendar and store it in document properties
function GetWeekInfo(baseDate = null) {
  const props = PropertiesService.getDocumentProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const infoSheet = ss.getSheetByName("INFO");

  const rawDate = infoSheet.getRange("D3").getValue();
  const startDate = baseDate instanceof Date ? new Date(baseDate) : new Date(rawDate);

  CURRENT_PERIOD_START = startDate;
  props.setProperty("CURRENT_PERIOD_START", startDate.toISOString());

  const periodNumber = infoSheet.getRange("D2").getValue();
  props.setProperty("CURRENT_PERIOD_NUMBER", periodNumber.toString());

  const weekList = {};
  for (let i = 0; i < 4; i++) {
    const weekStart = new Date(startDate);
    weekStart.setDate(weekStart.getDate() + (i * 7));

    const weekDates = [];
    for (let j = 0; j < 7; j++) {
      const d = new Date(weekStart);
      d.setDate(weekStart.getDate() + j);
      weekDates.push(Utilities.formatDate(d, Session.getScriptTimeZone(), "MM/dd/yyyy"));
    }

    weekList[`Week ${i + 1}`] = weekDates;
  }

  props.setProperty("weekList", JSON.stringify(weekList));

  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");
  const weeks = Object.values(weekList);
  let matchedIndex = weeks.findIndex(week => week.includes(todayStr));
  if (matchedIndex === -1) matchedIndex = 0;

  props.setProperty("currentWeekIndex", matchedIndex.toString());

  UpdateWeekView(matchedIndex);
}

// Move between previous or next week, regenerating the 4-week block if needed
function CheckDate(targetIndex) {
  const infoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("INFO");
  const props = PropertiesService.getDocumentProperties();

  const weekListRaw = props.getProperty("weekList");
  const weekList = weekListRaw ? JSON.parse(weekListRaw) : {};

  if (targetIndex >= 0 && targetIndex < 4) {
    props.setProperty("currentWeekIndex", targetIndex.toString());
    UpdateWeekView(targetIndex);
    return;
  }

  const currentPeriod = parseInt(infoSheet.getRange("D2").getValue(), 10) || 1;
  let direction;
  let periodStart;
  let period;

  if (targetIndex > 3) {
    const lastDate = new Date(weekList["Week 4"][6]);
    periodStart = new Date(lastDate);
    periodStart.setDate(lastDate.getDate() + 1);

    period = currentPeriod + 1;
    if (period > 13) period = 1;

    direction = "forward";
    props.setProperty("currentWeekIndex", "0");
  } else {
    const firstDate = new Date(weekList["Week 1"][0]);
    periodStart = new Date(firstDate);
    periodStart.setDate(firstDate.getDate() - 28);

    period = currentPeriod - 1;
    if (period < 1) period = 13;

    direction = "backward";
    props.setProperty("currentWeekIndex", "3");
  }

  infoSheet.getRange("D3").setValue(periodStart);
  infoSheet.getRange("D2").setValue(period);

  GetWeekInfo();
  UpdateInfoSheet();

  const newIndex = parseInt(props.getProperty("currentWeekIndex"), 10) || 0;
  UpdateWeekView(newIndex);

  SpreadsheetApp.getUi().alert(`Moved ${direction} to new 4-week period.\nCurrent Period: ${period}`);
}

// Called by buttons: next / previous
function NextWeek() {
  const index = parseInt(PropertiesService.getDocumentProperties().getProperty("currentWeekIndex"), 10) || 0;
  CheckDate(index + 1);
}

function PreviousWeek() {
  const index = parseInt(PropertiesService.getDocumentProperties().getProperty("currentWeekIndex"), 10) || 0;
  CheckDate(index - 1);
}

// Update current selected week display on the INFO sheet
function UpdateWeekView(index) {
  const props = PropertiesService.getDocumentProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const info = ss.getSheetByName("INFO");

  const weekListRaw = props.getProperty("weekList");
  if (!weekListRaw) return;

  const weekList = JSON.parse(weekListRaw);
  const weekKeys = Object.keys(weekList);
  const week = weekList[weekKeys[index]];

  if (!week) return;

  info.getRange("D5").setValue(week[0]);
  info.getRange("F5").setValue(week[6]);

  UpdateInfoSheet();
}

// Update INFO sheet with 4-week breakdown
function UpdateInfoSheet() {
  const props = PropertiesService.getDocumentProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const info = ss.getSheetByName("INFO");

  const weekListRaw = props.getProperty("weekList");
  if (!weekListRaw) return;

  const weekList = JSON.parse(weekListRaw);
  const weekKeys = Object.keys(weekList);

  weekKeys.forEach((label, i) => {
    const weekRange = weekList[label];
    info.getRange("B" + (7 + i)).setValue(label);
    info.getRange("D" + (7 + i)).setValue(weekRange[0]);
    info.getRange("E" + (7 + i)).setValue(":");
    info.getRange("F" + (7 + i)).setValue(weekRange[6]);
  });

  const periodStart = props.getProperty("CURRENT_PERIOD_START");
  const periodNumber = props.getProperty("CURRENT_PERIOD_NUMBER");

  if (periodStart) {
    info.getRange("D3").setValue(new Date(periodStart));
  }

  if (periodNumber) {
    info.getRange("D2").setValue(parseInt(periodNumber, 10));
  }
}
