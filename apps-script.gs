// ============================================================
//  PHYSIO 4 EVENTS — Google Apps Script
//
//  Setup:
//  1. Go to script.google.com → open your project
//  2. Replace ALL code with this file
//  3. Deploy → Manage deployments → pencil icon
//  4. Set Version to "New version" → Deploy
//  (URL stays the same, no need to update the HTML)
// ============================================================

var SHEET_ID   = '1AmVMKy81-l5XGaPDlnNM2V_v499Y7XjUvu5vrOB6oVE';
var SHEET_NAME = 'Dane';

function getSheet() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = setupSheet(ss);
  return sheet;
}

function doGet(e) {
  var cb = e.parameter.callback;
  try {
    var sheet = getSheet();
    var result;

    if (e.parameter.action === 'save') {
      var rev = (e.parameter.rev || '').split(',');
      var evt = (e.parameter.evt || '').split(',');
      for (var i = 0; i < 12; i++) {
        sheet.getRange(i + 2, 2).setValue(rev[i] !== '' ? parseFloat(rev[i]) : '');
        sheet.getRange(i + 2, 3).setValue(evt[i] !== '' ? parseInt(evt[i])   : '');
      }
      result = { status: 'ok' };
    } else {
      var rows = sheet.getRange(2, 1, 12, 3).getValues();
      result = rows.map(function(r) {
        return {
          rev: r[1] === '' ? null : r[1],
          evt: r[2] === '' ? null : r[2]
        };
      });
    }

    return cb ? jsonpOut(cb, result) : jsonOut(result);

  } catch (err) {
    var errObj = { error: err.message };
    return cb ? jsonpOut(cb, errObj) : jsonOut(errObj);
  }
}

function setupSheet(ss) {
  var sheet  = ss.insertSheet(SHEET_NAME);
  var months = [
    'Styczeń','Luty','Marzec','Kwiecień','Maj','Czerwiec',
    'Lipiec','Sierpień','Wrzesień','Październik','Listopad','Grudzień'
  ];
  sheet.getRange(1, 1, 1, 3).setValues([['Miesiąc','Przychód (PLN)','Liczba eventów']]);
  months.forEach(function(m, i) { sheet.getRange(i + 2, 1).setValue(m); });
  return sheet;
}

function jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonpOut(callback, data) {
  return ContentService
    .createTextOutput(callback + '(' + JSON.stringify(data) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
