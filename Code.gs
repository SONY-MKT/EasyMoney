/**
 * EasyMoney - Google Apps Script Backend API
 * 
 * Instructions:
 * 1. Open Google Sheets (https://sheets.google.com).
 * 2. Create a new Spreadsheet and name it "EasyMoney".
 * 3. Go to "Extensions" > "Apps Script".
 * 4. Replace all code in the editor with this script.
 * 5. Run the `setupEasyMoneySheet` function once to initialize all Sheets & headers.
 * 6. Click "Deploy" > "New deployment".
 * 7. Choose type "Web app".
 * 8. Set "Execute as": "Me" (your email) and "Who has access": "Anyone".
 * 9. Copy the generated Web App URL.
 * 10. Paste the URL into your EasyMoney App settings under "Google Sheet Sync".
 */

// Helper to get active sheet or create if not exists
function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#E8EEE9");
    }
  }
  return sheet;
}

// 1. Setup Spreadsheet Tabs and Headers
function migrateEntriesHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Entries");
  if (!sheet) return;

  const targetHeaders = [
    "id", "date", "month", "year", "type", "category", "description", "amount", "currency", "exchangeRate", "amountKHR", "amountUSD", "displayAmount", "paymentMethod", "memberId", "note", "createdAt", "updatedAt"
  ];
  
  const currentData = sheet.getDataRange().getValues();
  if (currentData.length === 0) {
    sheet.appendRow(targetHeaders);
    sheet.getRange(1, 1, 1, targetHeaders.length).setFontWeight("bold").setBackground("#E8EEE9");
    return;
  }
  
  const currentHeaders = currentData[0];
  const missingCols = ["currency", "exchangeRate", "amountKHR", "amountUSD", "displayAmount"].filter(h => currentHeaders.indexOf(h) === -1);
  
  if (missingCols.length > 0 || currentHeaders.join(",") !== targetHeaders.join(",")) {
    const newData = [];
    newData.push(targetHeaders);
    
    for (let i = 1; i < currentData.length; i++) {
        const oldRow = currentData[i];
        const newRow = [];
        const recordObj = {};
        for(let j=0; j<currentHeaders.length; j++) {
            recordObj[currentHeaders[j]] = oldRow[j];
        }
        
        let amount = Number(recordObj.amount) || 0;
        let currency = recordObj.currency;
        if(currency !== "៛" && currency !== "$") currency = "៛"; // Default to ៛ for old records if missing
        
        let exRate = Number(recordObj.exchangeRate) || 4000;
        recordObj.currency = currency;
        recordObj.exchangeRate = exRate;

        if (!recordObj.amountKHR && !recordObj.amountUSD) {
            if (currency === "៛") {
                recordObj.amountKHR = amount;
                recordObj.amountUSD = amount / exRate;
                recordObj.displayAmount = amount + " ៛";
            } else {
                recordObj.amountUSD = amount;
                recordObj.amountKHR = amount * exRate;
                recordObj.displayAmount = "$" + amount;
            }
        }
        
        for (let j=0; j<targetHeaders.length; j++) {
            newRow.push(recordObj[targetHeaders[j]] !== undefined ? recordObj[targetHeaders[j]] : "");
        }
        newData.push(newRow);
    }
    
    sheet.clear();
    sheet.getRange(1, 1, newData.length, targetHeaders.length).setValues(newData);
    sheet.getRange(1, 1, 1, targetHeaders.length).setFontWeight("bold").setBackground("#E8EEE9");
    
    try {
      formatEasyMoneySheets();
      createEntriesView();
      refreshEntriesView();
    } catch (e) {}
  }
}

function formatEasyMoneySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Entries");
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return;
  const headers = data[0];
  
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#E8EEE9");
  
  const columnsToHide = ["id", "amountKHR", "amountUSD", "createdAt", "updatedAt"];
  headers.forEach((h, i) => {
    if (columnsToHide.indexOf(h) !== -1) {
      sheet.hideColumns(i + 1);
    } else {
      sheet.showColumns(i + 1);
    }
  });
  
  const khrIdx = headers.indexOf("amountKHR") + 1;
  if (khrIdx > 0 && sheet.getMaxRows() > 1) sheet.getRange(2, khrIdx, sheet.getMaxRows() - 1, 1).setNumberFormat("#,##0");
  
  const usdIdx = headers.indexOf("amountUSD") + 1;
  if (usdIdx > 0 && sheet.getMaxRows() > 1) sheet.getRange(2, usdIdx, sheet.getMaxRows() - 1, 1).setNumberFormat("$#,##0.00");
  
  const rateIdx = headers.indexOf("exchangeRate") + 1;
  if (rateIdx > 0 && sheet.getMaxRows() > 1) sheet.getRange(2, rateIdx, sheet.getMaxRows() - 1, 1).setNumberFormat("#,##0");
  
  sheet.getDataRange().setWrap(true);
  sheet.getRange(1, 1, 1, headers.length).setWrap(false);
  try {
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
      const w = sheet.getColumnWidth(i);
      if (w < 90) sheet.setColumnWidth(i, 90);
    }
  } catch (e) {}
}

function createEntriesView() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let viewSheet = ss.getSheetByName("Entries_View");
  if (!viewSheet) {
    viewSheet = ss.insertSheet("Entries_View", 1);
  }
}

function refreshEntriesView() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const entriesSheet = ss.getSheetByName("Entries");
  if (!entriesSheet) return;
  let viewSheet = ss.getSheetByName("Entries_View");
  if (!viewSheet) return;
  
  const data = entriesSheet.getDataRange().getValues();
  if (data.length === 0) return;
  
  // Retrieve Family members map
  const membersSheet = ss.getSheetByName("FamilyMembers");
  const membersMap = {};
  if (membersSheet) {
    const memData = membersSheet.getDataRange().getValues();
    if (memData.length > 1) {
      const memHeaders = memData[0];
      const memIdIdx = memHeaders.indexOf("id");
      const memNameIdx = memHeaders.indexOf("name");
      if (memIdIdx !== -1 && memNameIdx !== -1) {
        for (let i = 1; i < memData.length; i++) {
          membersMap[memData[i][memIdIdx]] = memData[i][memNameIdx];
        }
      }
    }
  }

  const headers = data[0];
  const wantedHeaders = ["date", "month", "year", "type", "category", "description", "displayAmount", "currency", "paymentMethod", "memberId", "note"];
  const niceHeaders = ["ថ្ងៃខែ", "ខែ", "ឆ្នាំ", "ប្រភេទ", "ប្រភេទចំណាយ/ចំណូល", "ពិពណ៌នា", "ចំនួនលុយ", "រូបិយប័ណ្ណ", "វិធីបង់ប្រាក់", "សមាជិក", "កំណត់ចំណាំ"];
  
  const idIdx = headers.indexOf("id");
  const dateIdx = headers.indexOf("date");
  const amountIdx = headers.indexOf("amount");
  const currencyIdx = headers.indexOf("currency");
  
  const colIndices = wantedHeaders.map(wh => headers.indexOf(wh));
  
  const newViewData = [];
  newViewData.push(niceHeaders);
  
  const records = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[dateIdx] && !row[idIdx]) continue;
    
    // Format Display Amount & member name
    const newRow = colIndices.map((idx, index) => {
      let val = idx !== -1 ? row[idx] : "";
      const wh = wantedHeaders[index];
      
      if (wh === "memberId" && val) {
        val = membersMap[val] || val;
      }
      
      if (wh === "currency") {
        if (val === "$") val = "USD";
        if (val === "៛") val = "KHR";
      }
      
      if (wh === "displayAmount" && amountIdx !== -1 && currencyIdx !== -1) {
        const amount = Number(row[amountIdx]) || 0;
        const cur = row[currencyIdx] || "៛";
        if (cur === "$") {
          let parts = amount.toFixed(2).split(".");
          parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
          val = "$" + parts.join(".");
        } else {
          let str = Math.round(amount).toString();
          val = str.replace(/\B(?=(\d{3})+(?!\d))/g, ",") + " ៛";
        }
      }
      
      return val;
    });
    
    records.push({ rowData: newRow, date: row[dateIdx] || new Date(0) });
  }
  
  records.sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime());
  newViewData.push(...records.map(r => r.rowData));
  
  viewSheet.clear();
  if (newViewData.length > 0) {
    viewSheet.getRange(1, 1, newViewData.length, newViewData[0].length).setValues(newViewData);
  }
  
  viewSheet.setFrozenRows(1);
  const headerRange = viewSheet.getRange(1, 1, 1, newViewData[0].length);
  headerRange.setFontWeight("bold").setBackground("#d4edda");
  headerRange.setWrap(false); // disable wrap on headers to avoid weird splitting
  
  try {
    for (let i = 1; i <= newViewData[0].length; i++) {
      viewSheet.autoResizeColumn(i);
      const w = viewSheet.getColumnWidth(i);
      
      // Ensure minimum width for headers
      let minWidth = 80;
      if (niceHeaders[i-1] === "ថ្ងៃខែ") minWidth = 100;
      if (niceHeaders[i-1] === "ពិពណ៌នា") minWidth = 180;
      if (niceHeaders[i-1] === "ចំនួនលុយ") minWidth = 120;
      if (niceHeaders[i-1] === "កំណត់ចំណាំ") minWidth = 200;
      
      if (w < minWidth) viewSheet.setColumnWidth(i, minWidth);
    }
  } catch (e) {}

  if (newViewData.length > 1) {
    viewSheet.getRange(2, 1, newViewData.length - 1, newViewData[0].length).setWrap(true);
  }
  
  if (viewSheet.getFilter()) {
    viewSheet.getFilter().remove();
  }
  if (newViewData.length > 1) {
    viewSheet.getRange(1, 1, newViewData.length, newViewData[0].length).createFilter();
  }
  
  const rules = [];
  const typeIdx = niceHeaders.indexOf("ប្រភេទ") + 1;
  if (typeIdx > 0 && newViewData.length > 1) {
    const letter = String.fromCharCode(64 + typeIdx);
    const rangeStr = letter + "2:" + letter + newViewData.length;
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("ចំណូល").setBackground("#d4edda").setRanges([viewSheet.getRange(rangeStr)]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("ចំណាយ").setBackground("#f8d7da").setRanges([viewSheet.getRange(rangeStr)]).build());
    viewSheet.setConditionalFormatRules(rules);
  }
}

function setupEasyMoneySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Headers definitions
  const entriesHeaders = [
    "id", "date", "month", "year", "type", "category", "description", "amount", "currency", "exchangeRate", "amountKHR", "amountUSD", "displayAmount", "paymentMethod", "memberId", "note", "createdAt", "updatedAt"
  ];
  const membersHeaders = ["id", "name", "pin", "role", "isActive"];
  const categoriesHeaders = ["id", "type", "name"];
  const settingsHeaders = ["key", "value"];
  const logsHeaders = ["timestamp", "action", "user", "detail"];
  const importHistoryHeaders = ["id", "date", "importedCount", "status"];
  
  getOrCreateSheet("Entries", entriesHeaders);
  migrateEntriesHeaders();
  getOrCreateSheet("FamilyMembers", membersHeaders);
  getOrCreateSheet("Categories", categoriesHeaders);
  getOrCreateSheet("Settings", settingsHeaders);
  getOrCreateSheet("Logs", logsHeaders);
  getOrCreateSheet("ImportHistory", importHistoryHeaders);
  
  // Clean default Sheet1 if empty and other sheets exist
  const sheet1 = ss.getSheetByName("Sheet1");
  if (sheet1 && ss.getSheets().length > 1) {
    try {
      ss.deleteSheet(sheet1);
    } catch (e) {
      Logger.log("Could not delete Sheet1: " + e.message);
    }
  }
  
  // Seed initial Admin & Member if FamilyMembers empty
  const membersSheet = ss.getSheetByName("FamilyMembers");
  if (membersSheet.getLastRow() <= 1) {
    membersSheet.appendRow(["1", "Admin", "1234", "Admin", "true"]);
    membersSheet.appendRow(["2", "Member", "1111", "Member", "true"]);
  }
  
  // Seed default categories if Categories empty
  const catsSheet = ss.getSheetByName("Categories");
  if (catsSheet.getLastRow() <= 1) {
    const defaultCats = [
      ["c1", "ចំណូល", "ប្រាក់ខែ"],
      ["c2", "ចំណូល", "លក់ដូរ"],
      ["c3", "ចំណូល", "អាជីវកម្ម"],
      ["c4", "ចំណូល", "ជំនួយ"],
      ["c5", "ចំណូល", "ផ្សេងៗ"],
      ["c6", "ចំណាយ", "អាហារ"],
      ["c7", "ចំណាយ", "សាំង"],
      ["c8", "ចំណាយ", "ផ្ទះ"],
      ["c9", "ចំណាយ", "ទឹកភ្លើង"],
      ["c10", "ចំណាយ", "សុខភាព"],
      ["c11", "ចំណាយ", "កូនៗ"],
      ["c12", "ចំណាយ", "សាលារៀន"],
      ["c13", "ចំណាយ", "បំណុល"],
      ["c14", "ចំណាយ", "ទូរស័ព្ទ/Internet"],
      ["c15", "ចំណាយ", "ផ្សេងៗ"]
    ];
    defaultCats.forEach(row => catsSheet.appendRow(row));
  }
  
  try {
    formatEasyMoneySheets();
    createEntriesView();
    refreshEntriesView();
  } catch (e) {
    Logger.log("Formatting error: " + e.message);
  }
  
  return "Setup successfully completed!";
}

// Handler for CORS response
function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// 2. HTTP GET Endpoint (Read operations)
function doGet(e) {
  try {
    const action = e.parameter.action;
    
    // Auto setup on load if sheets are missing
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss.getSheetByName("Entries")) {
      setupEasyMoneySheet();
    }
    
    if (action === "test") {
      migrateEntriesHeaders();
      return jsonResponse({ status: "success", message: "ភ្ជាប់ Google Sheet បានជោគជ័យ", activeUser: Session.getActiveUser().getEmail() });
    }
    
    if (action === "getRecords" || !action) {
      const records = getRecords();
      return jsonResponse({ status: "success", records: records });
    }
    
    if (action === "getSettings") {
      const settings = getSettings();
      return jsonResponse({ status: "success", settings: settings });
    }

    if (action === "getUsers") {
      const users = getUsers();
      return jsonResponse({ status: "success", users: users });
    }

    if (action === "getCategories") {
      const categories = getCategories();
      return jsonResponse({ status: "success", categories: categories });
    }
    
    return jsonResponse({ status: "error", message: "រកមិនឃើញ Action ដែលបានស្នើសុំ" });
    
  } catch (err) {
    return jsonResponse({ status: "error", message: "កំហុស: " + err.toString() });
  }
}

// 3. HTTP POST Endpoint (Create, Update, Delete operations)
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const user = payload.user || "Unknown";
    
    // Auto setup if needed
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss.getSheetByName("Entries")) {
      setupEasyMoneySheet();
    }
    
    if (action === "test") {
      migrateEntriesHeaders();
      return jsonResponse({ status: "success", message: "ភ្ជាប់ Google Sheet បានជោគជ័យ" });
    }
    
    if (action === "addRecord") {
      const newRec = addRecord(payload.data);
      saveLog("ADD", user, "កត់ត្រាថ្មី ID: " + newRec.id + ", ប្រភេទ: " + newRec.type + ", ចំនួន: " + newRec.amount);
      return jsonResponse({ status: "success", record: newRec });
    }
    
    if (action === "updateRecord") {
      const updatedRec = updateRecord(payload.data);
      saveLog("UPDATE", user, "កែប្រែ ID: " + updatedRec.id + ", ប្រភេទ: " + updatedRec.type + ", ចំនួន: " + updatedRec.amount);
      return jsonResponse({ status: "success", record: updatedRec });
    }
    
    if (action === "deleteRecord") {
      deleteRecord(payload.id);
      saveLog("DELETE", user, "លុប ID: " + payload.id);
      return jsonResponse({ status: "success", message: "លុបបានជោគជ័យ" });
    }
    
    if (action === "syncAll") {
      // Bulk import or overwrite
      const syncStatus = syncAllLocalData(payload.data);
      saveLog("SYNC_ALL", user, "Sync ទិន្នន័យពី Local ទៅ Google Sheet ចំនួន " + payload.data.length + " ជួរ");
      return jsonResponse({ status: "success", message: "Sync ជោគជ័យ", count: payload.data.length });
    }
    
    if (action === "syncUsers") {
      const syncStatus = syncAllUsers(payload.data);
      saveLog("SYNC_USERS", user, "Sync អ្នកប្រើប្រាស់ពី Local ទៅ Google Sheet");
      return jsonResponse({ status: "success", message: "Sync អ្នកប្រើប្រាស់ជោគជ័យ", count: payload.data.length });
    }

    if (action === "syncCategories") {
      const syncStatus = syncAllCategories(payload.data);
      saveLog("SYNC_CATEGORIES", user, "Sync ប្រភេទពី Local ទៅ Google Sheet");
      return jsonResponse({ status: "success", message: "Sync ប្រភេទជោគជ័យ", count: payload.data.length });
    }

    if (action === "syncSettings") {
      saveSettings(payload.data);
      saveLog("SYNC_SETTINGS", user, "Sync ការកំណត់ (Settings) ទៅ Google Sheet");
      return jsonResponse({ status: "success", message: "Sync ការកំណត់ជោគជ័យ" });
    }
    
    return jsonResponse({ status: "error", message: "សកម្មភាពមិនត្រឹមត្រូវ" });
    
  } catch (err) {
    return jsonResponse({ status: "error", message: "កំហុស POST: " + err.toString() });
  }
}

// 4. Core functions implementations

// Retrieve all finance records
function getRecords() {
  migrateEntriesHeaders();
  const sheet = getOrCreateSheet("Entries");
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const list = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rec = {};
    headers.forEach((h, idx) => {
      let val = row[idx];
      // Convert specific columns to numbers
      if (["month", "year", "amount", "amountKHR", "amountUSD", "exchangeRate", "createdAt", "updatedAt"].indexOf(h) !== -1) {
        val = val !== "" ? Number(val) : 0;
      }
      
      // Fix date formatting for Google Sheets Date objects
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      
      rec[h] = val;
    });
    list.push(rec);
  }
  return list;
}

// Add a single record
function addRecord(recordData) {
  migrateEntriesHeaders();
  const sheet = getOrCreateSheet("Entries");
  
  // Generate random id if empty
  if (!recordData.id) {
    recordData.id = "gs_" + new Date().getTime() + "_" + Math.floor(Math.random() * 1000);
  }
  if (!recordData.createdAt) recordData.createdAt = Date.now();
  if (!recordData.updatedAt) recordData.updatedAt = Date.now();
  
  let currency = recordData.currency;
  if (currency !== "៛" && currency !== "$") currency = "៛";
  recordData.currency = currency;
  let exRate = Number(recordData.exchangeRate) || 4000;
  recordData.exchangeRate = exRate;
  let amount = Number(recordData.amount) || 0;
  if (!recordData.amountKHR && !recordData.amountUSD) {
     if (currency === "៛") {
         recordData.amountKHR = amount;
         recordData.amountUSD = amount / exRate;
         recordData.displayAmount = amount + " ៛";
     } else {
         recordData.amountUSD = amount;
         recordData.amountKHR = amount * exRate;
         recordData.displayAmount = "$" + amount;
     }
  }

  const headers = sheet.getDataRange().getValues()[0];
  const row = headers.map(h => recordData[h] !== undefined ? recordData[h] : "");
  
  sheet.appendRow(row);
  
  try {
    formatEasyMoneySheets();
    createEntriesView();
    refreshEntriesView();
  } catch (e) {}
  
  return recordData;
}

// Update an existing record
function updateRecord(recordData) {
  migrateEntriesHeaders();
  const sheet = getOrCreateSheet("Entries");
  const data = sheet.getDataRange().getValues();
  const idColIndex = data[0].indexOf("id");
  
  if (idColIndex === -1) throw new Error("id column not found in Google Sheet");
  
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idColIndex]) === String(recordData.id)) {
      rowIndex = i + 1; // 1-based index including headers
      break;
    }
  }
  
  const headers = data[0];
  if (!recordData.updatedAt) recordData.updatedAt = Date.now();
  
  let currency = recordData.currency;
  if (currency !== "៛" && currency !== "$") currency = "៛";
  recordData.currency = currency;
  let exRate = Number(recordData.exchangeRate) || 4000;
  recordData.exchangeRate = exRate;
  let amount = Number(recordData.amount) || 0;
  if (!recordData.amountKHR && !recordData.amountUSD) {
     if (currency === "៛") {
         recordData.amountKHR = amount;
         recordData.amountUSD = amount / exRate;
         recordData.displayAmount = amount + " ៛";
     } else {
         recordData.amountUSD = amount;
         recordData.amountKHR = amount * exRate;
         recordData.displayAmount = "$" + amount;
     }
  }

  const row = headers.map(h => recordData[h] !== undefined ? recordData[h] : "");
  
  if (rowIndex !== -1) {
    // Overwrite row
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
  } else {
    // If not found, add as new
    sheet.appendRow(row);
  }
  
  try {
    formatEasyMoneySheets();
    createEntriesView();
    refreshEntriesView();
  } catch (e) {}
  
  return recordData;
}

// Delete a single record matching ID
function deleteRecord(id) {
  const sheet = getOrCreateSheet("Entries");
  const data = sheet.getDataRange().getValues();
  const idColIndex = data[0].indexOf("id");
  
  if (idColIndex === -1) throw new Error("id column not found");
  
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][idColIndex]) === String(id)) {
      sheet.deleteRow(i + 1);
    }
  }
  
  try {
    formatEasyMoneySheets();
    createEntriesView();
    refreshEntriesView();
  } catch (e) {}
}

// Overwrite or update configuration settings in Google Sheet
function saveSettings(settingsObj) {
  const sheet = getOrCreateSheet("Settings");
  const headers = ["key", "value"];
  sheet.clear();
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#E8EEE9");
  
  if (settingsObj) {
    const rows = Object.keys(settingsObj).map(key => {
      let val = settingsObj[key];
      if (typeof val === 'object' && val !== null) {
        val = JSON.stringify(val);
      } else if (val === undefined || val === null) {
        val = "";
      } else {
        val = String(val);
      }
      return [key, val];
    });
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
  }
}

// Retrieve family users list from sheet
function getUsers() {
  const sheet = getOrCreateSheet("FamilyMembers");
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const item = {};
    headers.forEach((h, idx) => {
      let val = row[idx];
      if (h === 'isActive') {
        val = (val === true || val === 'true');
      }
      item[h] = val;
    });
    list.push(item);
  }
  return list;
}

// Retrieve categories list from sheet
function getCategories() {
  const sheet = getOrCreateSheet("Categories");
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const item = {};
    headers.forEach((h, idx) => {
      item[h] = row[idx];
    });
    list.push(item);
  }
  return list;
}

// Retrieve general sheet configuration settings
function getSettings() {
  const sheet = getOrCreateSheet("Settings");
  const data = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      settings[data[i][0]] = data[i][1];
    }
  }
  return settings;
}

// Bulk sync local data by flushing existing rows and re-populating
function syncAllLocalData(recordsArray) {
  const sheet = getOrCreateSheet("Entries");
  const headers = [
    "id", "date", "month", "year", "type", "category", "description", "amount", "currency", "exchangeRate", "amountKHR", "amountUSD", "displayAmount", "paymentMethod", "memberId", "note", "createdAt", "updatedAt"
  ];
  
  // Reset sheet with headers
  sheet.clear();
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#E8EEE9");
  
  if (recordsArray && recordsArray.length > 0) {
    const rows = recordsArray.map(rec => {
      // Ensure defaults for missing fields
      let currency = rec.currency;
      if (currency !== "៛" && currency !== "$") currency = "៛";
      rec.currency = currency;
      
      let exRate = Number(rec.exchangeRate) || 4000;
      rec.exchangeRate = exRate;
      
      let amount = Number(rec.amount) || 0;
      
      if (!rec.amountKHR && !rec.amountUSD) {
         if (currency === "៛") {
             rec.amountKHR = amount;
             rec.amountUSD = amount / exRate;
             rec.displayAmount = amount + " ៛";
         } else {
             rec.amountUSD = amount;
             rec.amountKHR = amount * exRate;
             rec.displayAmount = "$" + amount;
         }
      }

      return headers.map(h => rec[h] !== undefined ? rec[h] : "");
    });
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  
  try {
    formatEasyMoneySheets();
    createEntriesView();
    refreshEntriesView();
  } catch (e) {}
  
  return true;
}

// Bulk sync users
function syncAllUsers(usersArray) {
  const sheet = getOrCreateSheet("FamilyMembers");
  const headers = ["id", "name", "pin", "role", "isActive"];
  
  sheet.clear();
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#E8EEE9");
  
  if (usersArray && usersArray.length > 0) {
    const rows = usersArray.map(u => {
      return headers.map(h => u[h] !== undefined ? u[h] : "");
    });
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  return true;
}

// Bulk sync categories
function syncAllCategories(categoriesArray) {
  const sheet = getOrCreateSheet("Categories");
  const headers = ["id", "type", "name"];
  
  sheet.clear();
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#E8EEE9");
  
  if (categoriesArray && categoriesArray.length > 0) {
    const rows = categoriesArray.map(c => {
      return headers.map(h => c[h] !== undefined ? c[h] : "");
    });
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  return true;
}

// 5. Save Operational Logs
function saveLog(action, user, detail) {
  try {
    const sheet = getOrCreateSheet("Logs");
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    sheet.appendRow([timestamp, action, user, detail]);
  } catch (e) {
    Logger.log("Exception while saving log: " + e.message);
  }
}
