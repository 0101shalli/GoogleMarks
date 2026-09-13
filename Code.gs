/**
 * Googlemarks - Unified Google Sheets Add-on
 * Modules: Sheet Manager, Sheet Combiner, Sheet Access Control
 */

const PASSCODE_SHEET_NAME = "passcodes";
const AUTO_LOCK_TIMEOUT_MINUTES = 15; // Auto-lock timeout duration

/**
 * Automatically creates menu, updates passcode sheet visibility based on ownership, 
 * and re-locks all protected sheets when opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Googlemarks')
    .addItem('Open Googlemarks', 'showSidebar')
    .addToUi();

  // Manage visibility of passcodes sheet based on owner status
  enforcePasscodeSheetVisibility();

  // Enforce lock state every time the spreadsheet is opened/re-opened
  reLockAllProtectedSheets();
}

function showSidebar() {
  enforcePasscodeSheetVisibility();
  const html = HtmlService.createTemplateFromFile('Sidebar')
    .evaluate()
    .setTitle('Googlemarks')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getActiveSheetName() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}

/**
 * Checks if the current user is the owner of the spreadsheet.
 */
function isSheetOwner() {
  const ownerEmail = SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail();
  const activeEmail = Session.getActiveUser().getEmail();
  
  if (!activeEmail) return true; // Fallback for container execution
  return ownerEmail === activeEmail;
}

/**
 * Ensures the passcodes sheet is strictly visible ONLY to the owner.
 */
function enforcePasscodeSheetVisibility() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const passSheet = ss.getSheetByName(PASSCODE_SHEET_NAME);
  
  if (!passSheet) return;

  if (isSheetOwner()) {
    passSheet.showSheet(); // Visible to owner
  } else {
    passSheet.hideSheet(); // Hidden from all non-owners
  }
}

/* ==========================================
   PASSCODE SHEET MANAGEMENT & SECURITY
   ========================================== */

/**
 * Creates and protects the 'passcodes' sheet.
 */
function getOrCreatePasscodeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let passSheet = ss.getSheetByName(PASSCODE_SHEET_NAME);
  
  if (!passSheet) {
    passSheet = ss.insertSheet(PASSCODE_SHEET_NAME);
    passSheet.appendRow(["Sheet Name", "Passcode", "Date Generated", "Status", "Unlocked At"]);
    passSheet.getRange("A1:E1").setFontWeight("bold");
  }

  // Hide the sheet from non-owners without imposing strict range locks
  // that block Web App background updates
  enforcePasscodeSheetVisibility();

  return passSheet;
}

function generate5DigitPasscode() {
  return Math.floor(10000 + Math.random() * 90000).toString();
}

/* ==========================================
   MODULE 1: SHEET MANAGER
   ========================================== */

function getSheetsInfo() {
  enforcePasscodeSheetVisibility();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const protectedSheets = getProtectedSheetsList();
  
  // Filter out the system passcode sheet from the general list
  return sheets
    .filter(sheet => sheet.getName() !== PASSCODE_SHEET_NAME)
    .map(sheet => ({
      name: sheet.getName(),
      id: sheet.getSheetId(),
      isHidden: sheet.isSheetHidden(),
      isProtected: protectedSheets.includes(sheet.getName())
    }));
}

function activateSheet(sheetName) {
  if (sheetName === PASSCODE_SHEET_NAME && !isSheetOwner()) {
    return; // Block non-owners from activating passcode sheet
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.showSheet();
    sheet.activate();
  }
}

function bulkDeleteSheets(sheetNames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  sheetNames.forEach(name => {
    if (name === PASSCODE_SHEET_NAME) return; // Prevent deleting passcode sheet
    const sheet = ss.getSheetByName(name);
    if (sheet && ss.getSheets().length > 1) {
      ss.deleteSheet(sheet);
    }
  });
}

function bulkHideSheets(sheetNames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  sheetNames.forEach(name => {
    if (name === PASSCODE_SHEET_NAME) return;
    const sheet = ss.getSheetByName(name);
    if (sheet) sheet.hideSheet();
  });
}

function bulkUnhideSheets(sheetNames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  sheetNames.forEach(name => {
    if (name === PASSCODE_SHEET_NAME && !isSheetOwner()) return; // Prevent non-owner unhiding
    const sheet = ss.getSheetByName(name);
    if (sheet) sheet.showSheet();
  });
}

function duplicateSheet(sheetName) {
  if (sheetName === PASSCODE_SHEET_NAME) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    const newSheet = sheet.copyTo(ss);
    newSheet.setName(sheetName + ' (Copy)');
  }
}

/* ==========================================
   MODULE 2: SHEET COMBINER
   ========================================== */

function combineSheets(sheetNames, masterSheetName, addSourceColumn) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let masterSheet = ss.getSheetByName(masterSheetName);
  
  if (!masterSheet) {
    masterSheet = ss.insertSheet(masterSheetName);
  } else {
    masterSheet.clear();
  }

  let headers = [];
  const allRows = [];

  sheetNames.forEach(name => {
    if (name === PASSCODE_SHEET_NAME) return;
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    if (data.length > 0) {
      data[0].forEach(header => {
        if (header && !headers.includes(header)) {
          headers.push(header);
        }
      });
    }
  });

  if (headers.length === 0) return 'No data found in selected sheets.';

  const finalHeaders = addSourceColumn ? ['Source Sheet', ...headers] : [...headers];
  allRows.push(finalHeaders);

  sheetNames.forEach(name => {
    if (name === PASSCODE_SHEET_NAME) return;
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    const sheetHeaders = data[0];
    for (let r = 1; r < data.length; r++) {
      const rowData = data[r];
      const newRow = new Array(finalHeaders.length).fill('');
      
      if (addSourceColumn) newRow[0] = name;

      sheetHeaders.forEach((h, cIdx) => {
        const targetIdx = addSourceColumn ? headers.indexOf(h) + 1 : headers.indexOf(h);
        if (targetIdx !== -1) newRow[targetIdx] = rowData[cIdx];
      });
      allRows.push(newRow);
    }
  });

  masterSheet.getRange(1, 1, allRows.length, finalHeaders.length).setValues(allRows);
  return `Successfully combined ${sheetNames.length} sheets into "${masterSheetName}".`;
}

/* ==========================================
   MODULE 3: ACCESS CONTROL & RE-LOCK ENGINE
   ========================================== */

function getProtectableSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets()
    .map(s => s.getName())
    .filter(name => name !== PASSCODE_SHEET_NAME);
}

function lockSelectedSheets(selectedSheets) {
  if (!isSheetOwner()) {
    throw new Error("Access Denied: Only the Spreadsheet Owner can lock sheets.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const passcodeSheet = getOrCreatePasscodeSheet();
  const results = [];

  selectedSheets.forEach(sheetName => {
    if (sheetName === PASSCODE_SHEET_NAME) return;
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const passcode = generate5DigitPasscode();

    // Lock sheet content using Google Sheet Protection
    const existingProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    existingProtections.forEach(p => p.remove());

    const protection = sheet.protect().setDescription(`Passcode Protected: ${sheetName}`);
    protection.removeEditors(protection.getEditors());
    if (protection.canDomainEdit()) protection.setDomainEdit(false);

    // Record in passcodes sheet
    const data = passcodeSheet.getDataRange().getValues();
    let rowIndex = -1;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === sheetName) {
        rowIndex = i + 1;
        break;
      }
    }

    const timestamp = new Date();
    if (rowIndex > 0) {
      passcodeSheet.getRange(rowIndex, 2, 1, 4).setValues([[passcode, timestamp, "LOCKED", ""]]);
    } else {
      passcodeSheet.appendRow([sheetName, passcode, timestamp, "LOCKED", ""]);
    }

    results.push({ sheetName, passcode });
  });

  return { success: true, generated: results };
}

function verifyAndUnlockSheet(sheetName, inputPasscode) {
  const passcodeSheet = getOrCreatePasscodeSheet();
  const data = passcodeSheet.getDataRange().getValues();
  
  let storedPasscode = null;
  let rowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === sheetName) {
      storedPasscode = String(data[i][1]);
      rowIndex = i + 1;
      break;
    }
  }

  if (!storedPasscode) {
    return {
      success: false,
      message: `The sheet "${sheetName}" is not passcode protected or does not exist. Please check if you are on the right sheet.`
    };
  }

  if (String(inputPasscode).trim() !== storedPasscode) {
    return {
      success: false,
      message: `Incorrect passcode for "${sheetName}". Please enter the correct passcode or verify you are on the right sheet.`
    };
  }

  // Remove protection to unlock
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    protections.forEach(p => p.remove());
  }

  // Record UNLOCKED status and timestamp
  const now = new Date();
  passcodeSheet.getRange(rowIndex, 4).setValue("UNLOCKED");
  passcodeSheet.getRange(rowIndex, 5).setValue(now);

  ensureAutoLockTriggerExists();

  return {
    success: true,
    message: `Sheet "${sheetName}" unlocked for ${AUTO_LOCK_TIMEOUT_MINUTES} minutes!`
  };
}

function getProtectedSheetsList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const passSheet = ss.getSheetByName(PASSCODE_SHEET_NAME);
  if (!passSheet) return [];
  
  const data = passSheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) list.push(data[i][0]);
  }
  return list;
}

function enforcePasscodeSheetVisibility() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const passSheet = ss.getSheetByName(PASSCODE_SHEET_NAME);
  
  if (!passSheet) return;

  // Non-owners can only hide sheets they have rights to, or we let Google handles sheet hiding safely
  try {
    if (isSheetOwner()) {
      passSheet.showSheet();
    } else {
      passSheet.hideSheet();
    }
  } catch (e) {
    // Suppress permission errors on open for non-owners
  }
}

function reLockAllProtectedSheets() {
  // CRITICAL FIX: Non-owners MUST NOT attempt to re-lock or modify protections onOpen
  if (!isSheetOwner()) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const passcodeSheet = ss.getSheetByName(PASSCODE_SHEET_NAME);
  if (!passcodeSheet) return;

  const data = passcodeSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const sheetName = data[i][0];
    const sheet = ss.getSheetByName(sheetName);

    if (sheet) {
      const existingProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      existingProtections.forEach(p => p.remove());

      const protection = sheet.protect().setDescription(`Passcode Protected: ${sheetName}`);
      protection.removeEditors(protection.getEditors());
      if (protection.canDomainEdit()) protection.setDomainEdit(false);

      passcodeSheet.getRange(i + 1, 4).setValue("LOCKED");
      passcodeSheet.getRange(i + 1, 5).setValue("");
    }
  }
}

function checkAndLockExpiredSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const passcodeSheet = ss.getSheetByName(PASSCODE_SHEET_NAME);
  if (!passcodeSheet) return;

  const data = passcodeSheet.getDataRange().getValues();
  const now = new Date().getTime();
  let hasUnlockedSheetsRemaining = false;

  for (let i = 1; i < data.length; i++) {
    const sheetName = data[i][0];
    const status = data[i][3];
    const unlockedAtRaw = data[i][4];

    if (status === "UNLOCKED" && unlockedAtRaw) {
      const unlockedAt = new Date(unlockedAtRaw).getTime();
      const elapsedMinutes = (now - unlockedAt) / (1000 * 60);

      if (elapsedMinutes >= AUTO_LOCK_TIMEOUT_MINUTES) {
        const sheet = ss.getSheetByName(sheetName);
        if (sheet) {
          const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
          protections.forEach(p => p.remove());

          const protection = sheet.protect().setDescription(`Passcode Protected: ${sheetName}`);
          protection.removeEditors(protection.getEditors());
          if (protection.canDomainEdit()) protection.setDomainEdit(false);
        }

        passcodeSheet.getRange(i + 1, 4).setValue("LOCKED");
        passcodeSheet.getRange(i + 1, 5).setValue("");
      } else {
        hasUnlockedSheetsRemaining = true;
      }
    }
  }

  if (!hasUnlockedSheetsRemaining) {
    deleteAutoLockTriggers();
  }
}

function ensureAutoLockTriggerExists() {
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(t => t.getHandlerFunction() === 'checkAndLockExpiredSheets');
  
  if (!exists) {
    ScriptApp.newTrigger('checkAndLockExpiredSheets')
      .timeBased()
      .everyMinutes(1)
      .create();
  }
}

function deleteAutoLockTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'checkAndLockExpiredSheets') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

/* ==========================================
   WEB APP BRIDGE (PRIVILEGE ESCALATION)
   ========================================== */

/**
 * Runs as the OWNER when receiving post requests from non-owners.
 */
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const result = verifyAndUnlockSheet(params.sheetName, params.passcode);
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({success: false, message: err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Sends passcode requests from client sidebar to the Owner Web App.
 */
/**
 * Sends passcode requests from client sidebar to the Owner Web App.
 */
function proxyUnlockSheet(sheetName, passcode) {
  const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbwsW8TBm-WurrP7ONS1xjmjtMEZuY-6TnHHNVSGTQv-y_6Xi4OoHgKd0sv9uuEFrw8X/exec"; 
  
  // FIXED: Only block execution if the placeholder string hasn't been replaced or URL is empty
  if (!WEB_APP_URL || WEB_APP_URL === "PASTE_YOUR_WEB_APP_URL_HERE") {
    return { success: false, message: "Error: Web App URL is missing in Code.gs." };
  }

  const payload = { sheetName: sheetName, passcode: passcode };
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(WEB_APP_URL, options);
    const text = response.getContentText();
    
    // Catch Google HTML error pages (e.g., 403 Forbidden / Authorization required)
    if (text.startsWith("<!DOCTYPE") || text.includes("<html")) {
      return { 
        success: false, 
        message: "Web App Access Denied: Make sure Web App deployment access is set to 'Anyone'." 
      };
    }

    return JSON.parse(text);
  } catch (e) {
    return { success: false, message: "Bridge Error: " + e.message };
  }
}
