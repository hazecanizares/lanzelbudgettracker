// ============================================================
// BUDGET TRACKER — Google Apps Script
// Paste this entire file into script.google.com
// ============================================================

// ⚙️ STEP 1: Replace this with your Google Sheet ID
// (the long string in your Sheet URL between /d/ and /edit)
var SHEET_ID = 'PASTE_YOUR_SHEET_ID_HERE';

var SHEET_NAME = 'BudgetData'; // Do not change


// ============================================================
// doGet — called when the HTML app loads data
// ============================================================
function doGet(e) {
  var action = e && e.parameter && e.parameter.action;

  if (action === 'load') {
    return handleLoad();
  }

  // Default: return a simple status page so you can test the URL works
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'Budget Tracker script is running.' }))
    .setMimeType(ContentService.MimeType.JSON);
}


// ============================================================
// doPost — called when the HTML app saves data
// ============================================================
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.action === 'save') {
      return handleSave(body.data);
    }
    return jsonResponse({ status: 'error', message: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}


// ============================================================
// LOAD — read all data from the sheet and return as JSON
// ============================================================
function handleLoad() {
  try {
    var sheet = getOrCreateSheet();
    var range = sheet.getRange(1, 1, 1, 1);
    var raw = range.getValue();

    if (!raw || raw === '') {
      // No data yet — return an empty default structure
      return jsonResponse({
        status: 'ok',
        data: {
          expenses: [],
          banks: [
            { id: 'b1', name: 'BDO' },
            { id: 'b2', name: 'BPI' },
            { id: 'b3', name: 'GCash' }
          ],
          finEntries: { allan: [], hazel: [] },
          nextId: 1,
          nextFinId: 1
        }
      });
    }

    var data = JSON.parse(raw);
    return jsonResponse({ status: 'ok', data: data });

  } catch (err) {
    return jsonResponse({ status: 'error', message: 'Load failed: ' + err.toString() });
  }
}


// ============================================================
// SAVE — write all data to the sheet as JSON in cell A1
// ============================================================
function handleSave(data) {
  try {
    var sheet = getOrCreateSheet();

    // Store the full JSON blob in A1
    sheet.getRange(1, 1).setValue(JSON.stringify(data));

    // Also write a human-readable timestamp in B1
    sheet.getRange(1, 2).setValue('Last saved: ' + new Date().toLocaleString());

    // Also write expenses as readable rows starting from row 3
    // This makes the sheet readable without needing the app
    writeReadableExpenses(sheet, data);

    return jsonResponse({ status: 'ok', message: 'Saved successfully' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: 'Save failed: ' + err.toString() });
  }
}


// ============================================================
// READABLE ROWS — writes expenses as a human-readable table
// starting at row 3 so you can read the sheet like a spreadsheet
// ============================================================
function writeReadableExpenses(sheet, data) {
  try {
    // Clear old readable rows (row 3 downward)
    var lastRow = sheet.getLastRow();
    if (lastRow >= 3) {
      sheet.getRange(3, 1, lastRow - 2, 10).clearContent();
    }

    // Header row at row 3
    var headers = ['Month', 'Year', 'Category', 'Description', 'Bank', 'Amount', 'Paid By', 'Paid?', 'Installment', 'Months Left'];
    sheet.getRange(3, 1, 1, headers.length).setValues([headers]);

    var expenses = data.expenses || [];
    if (expenses.length === 0) return;

    var MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    var rows = expenses.map(function(e) {
      return [
        MONTHS[e.month] || e.month,
        e.year,
        e.category || '',
        e.desc || '',
        e.bank || '',
        e.amount || 0,
        e.paidBy || '',
        e.paid ? 'Yes' : 'No',
        e.installment === 'yes' ? 'Yes' : 'No',
        e.installment === 'yes' ? (e.installmentMonths || 0) : ''
      ];
    });

    sheet.getRange(4, 1, rows.length, headers.length).setValues(rows);

  } catch (err) {
    // Non-fatal — don't break the save if readable rows fail
    Logger.log('writeReadableExpenses error: ' + err.toString());
  }
}


// ============================================================
// HELPERS
// ============================================================
function getOrCreateSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
