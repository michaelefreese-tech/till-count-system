// security.gs

function RefreshMenus() {
  try {
    SpreadsheetApp.getUi().createMenu("🔒 Security").addToUi();
    SpreadsheetApp.getUi().createMenu("🛠 Script Tools").addToUi();
  } catch (err) {
    // ignore placeholder menu errors
  }

  BuildMenus();

  if (!IsFileInitialized()) {
    BuildSetupMenu();
  }
}

function ApplySheetProtections() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();
  const me = Session.getEffectiveUser().getEmail();

  // Remove old sheet protections first
  ss.getSheets().forEach(sheet => {
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    protections.forEach(protection => {
      try {
        protection.remove();
      } catch (err) {
        // ignore protections we cannot remove
      }
    });
  });

  const lockSheet = (sheetName, description) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const protection = sheet.protect().setDescription(description);
    protection.setWarningOnly(false);

    const allowedCells = Info[sheetName]?.ALLOWEDCELLS || [];
    if (allowedCells.length > 0) {
      const unprotectedRanges = allowedCells.map(a1 => sheet.getRange(a1));
      protection.setUnprotectedRanges(unprotectedRanges);
    }

    try {
      protection.addEditor(me);
      protection.removeEditors(
        protection.getEditors().filter(user => user.getEmail() !== me)
      );
      if (protection.canDomainEdit()) protection.setDomainEdit(false);
    } catch (err) {
      // ignore editor cleanup issues
    }
  };

  REQUIRED_SHEETS.forEach(sheetName => {
    lockSheet(sheetName, `LOCKED_${sheetName}`);
  });

  props.setProperty("docLocked", "true");
}

function RemoveSheetProtections() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();

  ss.getSheets().forEach(sheet => {
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    protections.forEach(protection => {
      try {
        protection.remove();
      } catch (err) {
        // ignore protections we cannot remove
      }
    });
  });

  props.setProperty("docLocked", "false");
}

// Lock the document
function Lock() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();
  const status = props.getProperty("docLocked") === "true";

  if (status) {
    ui.alert("Document is already locked.");
    RefreshMenus();
    return;
  }

  ApplySheetProtections();
  props.deleteProperty("justUnlocked");

  RefreshMenus();
  ui.alert("Document is now locked.");
}

// Unlock with password prompt
function Unlock() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();
  const status = props.getProperty("docLocked") === "true";

  if (!status) {
    ui.alert("Document is already unlocked.");
    RefreshMenus();
    return;
  }

  const response = ui.prompt("Enter password to unlock:");
  if (response.getSelectedButton() !== ui.Button.OK) {
    RefreshMenus();
    return;
  }

  const password = response.getResponseText().trim();
  const stored = JSON.parse(props.getProperty("PASSWORD_DATA") || "null");

  if (!stored) {
    ui.alert("No password set. Please run SetupPassword() first.");
    RefreshMenus();
    return;
  }

  if (password === stored.password) {
    RemoveSheetProtections();
    props.setProperty("justUnlocked", "true");

    RefreshMenus();
    ui.alert("Document is now unlocked.");
  } else {
    const recovery = ui.prompt("Incorrect password.\nType 'recover' to reset using security questions, or Cancel.");

    if (
      recovery.getSelectedButton() === ui.Button.OK &&
      recovery.getResponseText().trim().toLowerCase() === "recover"
    ) {
      RecoverPassword();
    } else {
      ui.alert("Access denied.");
      RefreshMenus();
    }
  }
}

// Setup password and security questions
function SetupPassword() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();
  const stored = JSON.parse(props.getProperty("PASSWORD_DATA") || "null");

  if (stored) {
    const current = ui.prompt("Enter current password to change:");
    if (
      current.getSelectedButton() !== ui.Button.OK ||
      current.getResponseText().trim() !== stored.password
    ) {
      ui.alert("Incorrect password. Cannot change settings.");
      RefreshMenus();
      return;
    }
  }

  const passwordPrompt = ui.prompt("Set a new password:");
  if (passwordPrompt.getSelectedButton() !== ui.Button.OK) {
    RefreshMenus();
    return;
  }
  const password = passwordPrompt.getResponseText().trim();

  const q1Prompt = ui.prompt("Security Question 1 (e.g., Favorite color?)");
  if (q1Prompt.getSelectedButton() !== ui.Button.OK) {
    RefreshMenus();
    return;
  }
  const q1 = q1Prompt.getResponseText().trim();

  const a1Prompt = ui.prompt("Answer to Question 1:");
  if (a1Prompt.getSelectedButton() !== ui.Button.OK) {
    RefreshMenus();
    return;
  }
  const a1 = a1Prompt.getResponseText().trim().toLowerCase();

  const q2Prompt = ui.prompt("Security Question 2 (e.g., First pet's name?)");
  if (q2Prompt.getSelectedButton() !== ui.Button.OK) {
    RefreshMenus();
    return;
  }
  const q2 = q2Prompt.getResponseText().trim();

  const a2Prompt = ui.prompt("Answer to Question 2:");
  if (a2Prompt.getSelectedButton() !== ui.Button.OK) {
    RefreshMenus();
    return;
  }
  const a2 = a2Prompt.getResponseText().trim().toLowerCase();

  const data = { password, q1, a1, q2, a2 };
  props.setProperty("PASSWORD_DATA", JSON.stringify(data));

  RefreshMenus();
  ui.alert("Password and security questions set.");
}

// Recover password via Q&A
function RecoverPassword() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();
  const stored = JSON.parse(props.getProperty("PASSWORD_DATA") || "null");

  if (!stored) {
    ui.alert("No recovery data found.");
    RefreshMenus();
    return;
  }

  const a1Prompt = ui.prompt(stored.q1);
  if (a1Prompt.getSelectedButton() !== ui.Button.OK) {
    RefreshMenus();
    return;
  }
  const a1 = a1Prompt.getResponseText().trim().toLowerCase();

  const a2Prompt = ui.prompt(stored.q2);
  if (a2Prompt.getSelectedButton() !== ui.Button.OK) {
    RefreshMenus();
    return;
  }
  const a2 = a2Prompt.getResponseText().trim().toLowerCase();

  if (a1 === stored.a1 && a2 === stored.a2) {
    const newPassPrompt = ui.prompt("Enter your new password:");
    if (newPassPrompt.getSelectedButton() !== ui.Button.OK) {
      RefreshMenus();
      return;
    }

    stored.password = newPassPrompt.getResponseText().trim();
    props.setProperty("PASSWORD_DATA", JSON.stringify(stored));
    ui.alert("Password successfully reset.");
  } else {
    ui.alert("Security answers did not match. Cannot reset.");
  }

  RefreshMenus();
}

// Admin tool: Clear all password and recovery data
function ClearPassword() {
  PropertiesService.getDocumentProperties().deleteProperty("PASSWORD_DATA");
  RefreshMenus();
  SpreadsheetApp.getUi().alert("All password data cleared.");
}

// Show current script version
function ShowVersion() {
  const version = PropertiesService.getDocumentProperties().getProperty("SCRIPT_VERSION") || "Not set";
  SpreadsheetApp.getUi().alert("Current Version: " + version);
}

// Manually add version + changelog
function PromptNewVersion() {
  const ui = SpreadsheetApp.getUi();

  const versionPrompt = ui.prompt("Enter new version:");
  if (versionPrompt.getSelectedButton() !== ui.Button.OK) {
    RefreshMenus();
    return;
  }
  const newVersion = versionPrompt.getResponseText().trim();

  const logPrompt = ui.prompt("Enter changelog for version " + newVersion + ":");
  if (logPrompt.getSelectedButton() !== ui.Button.OK) {
    RefreshMenus();
    return;
  }
  const changelog = logPrompt.getResponseText();

  UpdateVersionWithLog(newVersion, changelog);
  RefreshMenus();
}

// Store changelog data
function UpdateVersionWithLog(version, changelog) {
  const props = PropertiesService.getDocumentProperties();
  const rawLog = props.getProperty("SCRIPT_CHANGELOG_DICT") || "[]";
  const log = JSON.parse(rawLog);
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM.dd.yyyy");

  log.push({ date: timestamp, version, log: changelog });

  props.setProperty("SCRIPT_CHANGELOG_DICT", JSON.stringify(log));
  props.setProperty("SCRIPT_VERSION", version);
}

// Display changelog history
function ViewChangeLog() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();
  const rawLog = props.getProperty("SCRIPT_CHANGELOG_DICT") || "[]";
  const logList = JSON.parse(rawLog);

  if (logList.length === 0) {
    ui.alert("No changelog available.");
    return;
  }

  let message = "Version History:\n\n";
  logList.forEach(entry => {
    message += `${entry.date} : ${entry.version} : ${entry.log}\n`;
  });

  ui.alert(message);
}

// Admin: Clear version history
function ClearVersionHistory() {
  const ui = SpreadsheetApp.getUi();
  const confirmation = ui.alert("⚠️ WARNING", "This will permanently delete version history. Proceed?", ui.ButtonSet.YES_NO);
  if (confirmation !== ui.Button.YES) {
    RefreshMenus();
    return;
  }

  const props = PropertiesService.getDocumentProperties();
  props.deleteProperty("SCRIPT_VERSION");
  props.deleteProperty("SCRIPT_CHANGELOG_DICT");

  RefreshMenus();
  ui.alert("✅ All version history cleared.");
}
