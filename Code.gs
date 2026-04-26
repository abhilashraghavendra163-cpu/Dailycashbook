// Code.gs
// DAILY CASHBOOK & DENOMINATION TRACKER - GOOGLE SHEETS BACKEND

const SHEET_NAMES = {
  SUMMARY: 'Daily_Summary',
  TRANSACTIONS: 'Transactions',
  DENOMINATIONS: 'Denominations',
  USERS: 'Users',
  BANK_TRANSACTIONS: 'Bank_Transactions',
  DAILY_REPORTS: 'Daily_Reports_Formatted'
};

// ============================================================
// RUN repairSystem() ONCE FROM EDITOR AFTER PASTING THIS CODE
// ============================================================

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.values(SHEET_NAMES).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    repairSheetHeaders(sheet, name);
  });
}

function repairSheetHeaders(sheet, name) {
  if (!sheet) return;
  let headers = [];
  if (name === SHEET_NAMES.SUMMARY) {
    headers = ['Date','Opening_Balance','Total_Deposits','Total_Expenses','Closing_Balance',
               'Bank_Opening_Balance','Bank_Total_Credits','Bank_Total_Debits','Bank_Closing_Balance',
               'Verified','Variance','Remarks','Counted_By','Verified_By'];
  } else if (name === SHEET_NAMES.TRANSACTIONS) {
    headers = ['ID','Date','Type','Mode','Name','Category','Description','Amount','Entered_by','Entry_Date','Entry_Time'];
  } else if (name === SHEET_NAMES.DENOMINATIONS) {
    headers = ['Date','500','200','100','50','20','10','5','2','1',' ','TOTAL'];
  } else if (name === SHEET_NAMES.USERS) {
    headers = ['Username','Password','Role'];
  } else if (name === SHEET_NAMES.BANK_TRANSACTIONS) {
    headers = ['ID','Date','Type','Mode','Reference','Description','Amount','Entered_by','Entry_Date','Entry_Time'];
  } else if (name === SHEET_NAMES.DAILY_REPORTS) {
    headers = ['Date','Summary_Type','Opening_Balance','Total_Inflow','Total_Outflow','Closing_Balance','Status','Generated_At'];
  }

  if (headers.length > 0) {
    // Always force-write correct headers to row 1
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#F3F4F6');
    sheet.setFrozenRows(1);

    // Add default admin user if Users sheet is empty
    if (name === SHEET_NAMES.USERS && sheet.getLastRow() <= 1) {
      sheet.appendRow(['admin', '123', 'Admin']);
    }
  }
}

function repairSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Force timezone to IST so dates are never shifted
  ss.setSpreadsheetTimeZone('Asia/Kolkata');
  
  setupSheets();

  // Delete corrupted sheets like Day_undefined
  ss.getSheets().forEach(s => {
    const n = s.getName();
    if (n.includes('undefined') || n.includes('null')) {
      try { ss.deleteSheet(s); } catch(e) {}
    }
  });

  Logger.log('System repair complete. Timezone set to IST. All headers fixed.');
}

// ============================================================

function doGet() {
  return HtmlService.createHtmlOutput('<h2>Cashbook API is running.</h2>')
    .setTitle('Cashbook API');
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const payload = data.payload;
    const date = data.date;

    let result;
    switch (action) {

      // USER MANAGEMENT
      case 'getUsers':
        result = getSheetData(SHEET_NAMES.USERS);
        break;
      case 'addUser':
        result = appendRow(SHEET_NAMES.USERS, payload);
        break;
      case 'editUser':
        result = editUser(payload);
        break;
      case 'deleteUser':
        result = deleteRow(SHEET_NAMES.USERS, 'Username', payload.Username);
        break;

      // SUMMARY / DAY DATA
      case 'getSummary':
        result = getSheetData(SHEET_NAMES.SUMMARY);
        break;
      case 'loadDayData':
        result = {
          summaries: getSheetData(SHEET_NAMES.SUMMARY),
          txs:       getSheetData(SHEET_NAMES.TRANSACTIONS),
          denoms:    getSheetData(SHEET_NAMES.DENOMINATIONS),
          bankTxs:   getSheetData(SHEET_NAMES.BANK_TRANSACTIONS)
        };
        break;
      case 'saveSummary':
        result = updateOrAppendSummary(payload);
        break;

      // CASH TRANSACTIONS
      case 'saveTransaction':
        result = appendRow(SHEET_NAMES.TRANSACTIONS, payload);
        break;
      case 'deleteTransaction':
        result = deleteRow(SHEET_NAMES.TRANSACTIONS, 'ID', payload.ID);
        break;
      case 'getTransactions':
        result = getSheetData(SHEET_NAMES.TRANSACTIONS);
        break;

      // BANK TRANSACTIONS
      case 'saveBankTransaction':
        result = appendRow(SHEET_NAMES.BANK_TRANSACTIONS, payload);
        break;
      case 'deleteBankTransaction':
        result = deleteRow(SHEET_NAMES.BANK_TRANSACTIONS, 'ID', payload.ID);
        break;
      case 'getBankTransactions':
        result = getSheetData(SHEET_NAMES.BANK_TRANSACTIONS);
        break;

      // DENOMINATIONS
      case 'saveDenominations':
        result = saveDenominations(payload, date);
        break;
      case 'getDenominations':
        result = getSheetData(SHEET_NAMES.DENOMINATIONS);
        break;

      // SEARCH
      case 'globalSearch':
        result = globalSearch(payload ? payload.query : '');
        break;

      default:
        result = { status: 'error', message: 'Unknown action: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// HELPER FUNCTIONS
// ============================================================

function getSheetData(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  // Normalize headers so column name variations don't break the app
  const headers = data[0].map(h => {
    const s = String(h).trim().toLowerCase().replace(/[\s_]/g, '');
    if (s === 'date')                                          return 'Date';
    if (s === 'amount')                                        return 'Amount';
    if (s === 'id')                                            return 'ID';
    if (s === 'type')                                          return 'Type';
    if (s === 'mode')                                          return 'Mode';
    if (s === 'name')                                          return 'Name';
    if (s === 'category')                                      return 'Category';
    if (s === 'description')                                   return 'Description';
    if (s === 'reference')                                     return 'Reference';
    if (s === 'openingbalance')                                return 'Opening_Balance';
    if (s === 'closingbalance')                                return 'Closing_Balance';
    if (s === 'verified')                                      return 'Verified';
    if (s === 'totaldeposits' || s === 'totalinflow')          return 'Total_Deposits';
    if (s === 'totalexpenses' || s === 'totaloutflow')         return 'Total_Expenses';
    if (s === 'bankopeningbalance')                            return 'Bank_Opening_Balance';
    if (s === 'bankclosingbalance')                            return 'Bank_Closing_Balance';
    if (s === 'banktotalcredits')                              return 'Bank_Total_Credits';
    if (s === 'banktotaldebits')                               return 'Bank_Total_Debits';
    if (s === 'username' || s === 'user')                      return 'Username';
    if (s === 'password' || s === 'pass')                      return 'Password';
    if (s === 'role')                                          return 'Role';
    if (s === 'variance')                                      return 'Variance';
    if (s === 'remarks')                                       return 'Remarks';
    if (s === 'countedby')                                     return 'Counted_By';
    if (s === 'verifiedby')                                    return 'Verified_By';
    if (s === 'enteredby')                                     return 'Entered_by';
    if (s === 'entrydate')                                     return 'Entry_Date';
    if (s === 'entrytime')                                     return 'Entry_Time';
    return String(h).trim().replace(/\s+/g, '_');
  });

  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { if (h && h.trim()) obj[h] = safeValue(row[i]); });
    return obj;
  }).filter(row => Object.values(row).some(v => v !== '' && v !== null && v !== undefined));
}

function appendRow(name, payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) { setupSheets(); sheet = ss.getSheetByName(name); }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => {
    const normalized = String(h).trim();
    return payload[normalized] !== undefined ? payload[normalized] : '';
  });
  sheet.appendRow(row);
  return { status: 'success' };
}

function deleteRow(name, key, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) return { status: 'error', message: 'Sheet not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const colIndex = headers.indexOf(key);
  if (colIndex === -1) return { status: 'error', message: 'Key column not found: ' + key };
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][colIndex]).trim() === String(value).trim()) {
      sheet.deleteRow(i + 1);
      return { status: 'success' };
    }
  }
  return { status: 'not_found' };
}

function editUser(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
  if (!sheet) return { status: 'error', message: 'Users sheet not found' };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === String(payload.Username).trim().toLowerCase()) {
      if (payload.Password !== undefined) sheet.getRange(i + 1, 2).setValue(payload.Password);
      if (payload.Role !== undefined)     sheet.getRange(i + 1, 3).setValue(payload.Role);
      return { status: 'success' };
    }
  }
  return { status: 'error', message: 'User not found' };
}

function saveDenominations(payload, date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.DENOMINATIONS);
  if (!sheet) { setupSheets(); sheet = ss.getSheetByName(SHEET_NAMES.DENOMINATIONS); }

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const dateColIndex = headers.indexOf('Date');
  let foundIndex = -1;

  for (let i = 1; i < data.length; i++) {
    if (isSameDay(data[i][dateColIndex], date) || isSameDay(data[i][dateColIndex], payload.Date)) {
      foundIndex = i;
      break;
    }
  }

  const rowData = headers.map(h => payload[h] !== undefined ? payload[h] : '');
  if (foundIndex > -1) {
    sheet.getRange(foundIndex + 1, 1, 1, headers.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  return { status: 'success' };
}

function updateOrAppendSummary(payload) {
  if (!payload.Date || String(payload.Date).includes('undefined')) {
    return { status: 'error', message: 'Invalid date in payload' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.SUMMARY);
  if (!sheet) { setupSheets(); sheet = ss.getSheetByName(SHEET_NAMES.SUMMARY); }

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const dateColIndex = headers.indexOf('Date');

  for (let i = 1; i < data.length; i++) {
    if (isSameDay(data[i][dateColIndex], payload.Date)) {
      const rowData = headers.map((h, hi) =>
        payload[h] !== undefined ? payload[h] : data[i][hi]
      );
      sheet.getRange(i + 1, 1, 1, headers.length).setValues([rowData]);
      if (payload.Verified) { try { drawDailySheet(payload.Date); } catch(e) {} }
      return { status: 'success' };
    }
  }

  const rowData = headers.map(h => payload[h] !== undefined ? payload[h] : '');
  sheet.appendRow(rowData);
  if (payload.Verified) { try { drawDailySheet(payload.Date); } catch(e) {} }
  return { status: 'success' };
}

// ============================================================
// DATE HELPERS
// ============================================================

function safeValue(val) {
  if (val === null || val === undefined || val === '') return '';
  if (typeof val === 'number' && val === 0) return 0;
  if (!val && val !== 0) return '';
  if (val instanceof Date) {
    // Always apply IST offset (UTC+5:30) to avoid date shifting on UTC servers
    const ist = new Date(val.getTime() + (5.5 * 60 * 60 * 1000));
    const y = ist.getUTCFullYear();
    const m = String(ist.getUTCMonth() + 1).padStart(2, '0');
    const d = String(ist.getUTCDate()).padStart(2, '0');
    return y + '-' + m + '-' + d;
  }
  return val;
}

function parseDateToParts(val) {
  if (!val) return null;
  if (val instanceof Date) {
    return { y: val.getFullYear(), m: val.getMonth() + 1, d: val.getDate() };
  }
  const s = String(val).trim();
  if (!s) return null;
  let match = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (match) return { y: parseInt(match[1]), m: parseInt(match[2]), d: parseInt(match[3]) };
  match = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
  if (match) return { y: parseInt(match[3]), m: parseInt(match[2]), d: parseInt(match[1]) };
  const dt = new Date(s);
  if (!isNaN(dt.getTime())) return { y: dt.getFullYear(), m: dt.getMonth() + 1, d: dt.getDate() };
  return null;
}

function isSameDay(d1, d2) {
  const p1 = parseDateToParts(d1);
  const p2 = parseDateToParts(d2);
  return p1 && p2 && p1.y === p2.y && p1.m === p2.m && p1.d === p2.d;
}

// ============================================================
// DRAW DAILY SHEET (called on Verify & Lock)
// ============================================================

function drawDailySheet(dateStr) {
  if (!dateStr || String(dateStr).includes('undefined')) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Day_' + dateStr;
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  else sheet.clear();

  const summaries = getSheetData(SHEET_NAMES.SUMMARY);
  const summary   = summaries.find(s => isSameDay(s.Date, dateStr)) || {};
  const allTxs    = getSheetData(SHEET_NAMES.TRANSACTIONS).filter(t => isSameDay(t.Date, dateStr));
  const deps      = allTxs.filter(t => t.Type === 'Deposit');
  const exps      = allTxs.filter(t => t.Type === 'Expense');
  const denoms    = getSheetData(SHEET_NAMES.DENOMINATIONS).filter(d => isSameDay(d.Date, dateStr));
  const bankTxs   = getSheetData(SHEET_NAMES.BANK_TRANSACTIONS).filter(t => isSameDay(t.Date, dateStr));
  const bCreds    = bankTxs.filter(t => t.Type === 'Credit');
  const bDebs     = bankTxs.filter(t => t.Type === 'Debit');

  const view = [];
  view.push(['DAILY CASHBOOK LEDGER - ' + dateStr, '', '', '']);
  view.push([]);
  view.push(['Status:', summary.Verified ? 'Verified & Locked' : 'Draft', '', '']);
  view.push(['Opening Balance:', summary.Opening_Balance || 0, '', '']);
  view.push([]);
  view.push(['--- RECEIPTS ---', 'Amount', '--- PAYMENTS ---', 'Amount']);

  const maxRows = Math.max(deps.length, exps.length, 1);
  for (let i = 0; i < maxRows; i++) {
    view.push([
      deps[i] ? deps[i].Category + ' (' + deps[i].Mode + ')' : '',
      deps[i] ? deps[i].Amount : '',
      exps[i] ? (exps[i].Name || '') + ': ' + exps[i].Category + ' (' + exps[i].Mode + ')' : '',
      exps[i] ? exps[i].Amount : ''
    ]);
  }

  view.push([]);
  view.push(['Total Receipts:', summary.Total_Deposits || 0, 'Total Payments:', summary.Total_Expenses || 0]);
  view.push(['Closing Balance:', summary.Closing_Balance || 0, '', '']);
  view.push([]);
  view.push(['--- BANK RECEIPTS ---', 'Amount', '--- BANK PAYMENTS ---', 'Amount']);

  const maxBank = Math.max(bCreds.length, bDebs.length, 1);
  for (let i = 0; i < maxBank; i++) {
    view.push([
      bCreds[i] ? bCreds[i].Description + ' (Ref: ' + (bCreds[i].Reference || 'N/A') + ')' : '',
      bCreds[i] ? bCreds[i].Amount : '',
      bDebs[i]  ? bDebs[i].Description  + ' (Ref: ' + (bDebs[i].Reference  || 'N/A') + ')' : '',
      bDebs[i]  ? bDebs[i].Amount  : ''
    ]);
  }

  view.push(['Bank Total Credits:', summary.Bank_Total_Credits || 0, 'Bank Total Debits:', summary.Bank_Total_Debits || 0]);
  view.push(['Bank Opening:', summary.Bank_Opening_Balance || 0, 'Bank Closing:', summary.Bank_Closing_Balance || 0]);
  view.push([]);

  if (denoms.length > 0) {
    view.push(['--- PHYSICAL CASH COUNT ---', '', '', '']);
    const d = denoms[0];
    [500, 200, 100, 50, 20, 10, 5, 2, 1].forEach(note => {
      const count = Number(d[note]) || 0;
      if (count > 0) view.push(['Note: Rs.' + note, count + ' notes', 'Value:', note * count]);
    });
    view.push([]);
    view.push(['Variance:', summary.Variance || 0, '', '']);
    view.push(['Remarks:', summary.Remarks || 'None', '', '']);
  }

  view.push([]);
  view.push(['Verified By:', summary.Verified_By || 'System', '', '']);
  view.push(['Report Generated At:', new Date().toLocaleString(), '', '']);

  sheet.getRange(1, 1, view.length, 4).setValues(view);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setFontSize(14).setBackground('#4F46E5').setFontColor('white');
  sheet.autoResizeColumns(1, 4);
  sheet.setFrozenRows(1);
}

// ============================================================
// GLOBAL SEARCH
// ============================================================

function globalSearch(query) {
  if (!query) return [];
  const q = String(query).toLowerCase().trim();
  const results = [];

  const searchIn = (sheetName, label) => {
    getSheetData(sheetName).forEach(row => {
      if (Object.values(row).join(' ').toLowerCase().includes(q)) {
        results.push(Object.assign({}, row, { _source: label }));
      }
    });
  };

  searchIn(SHEET_NAMES.TRANSACTIONS, 'Cash');
  searchIn(SHEET_NAMES.BANK_TRANSACTIONS, 'Bank');
  searchIn(SHEET_NAMES.SUMMARY, 'Summary');
  return results.slice(0, 100);
}
