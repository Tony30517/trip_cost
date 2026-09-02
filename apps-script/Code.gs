// ==== 設定：換其他行程時只要改這裡 ====
var FIREBASE_PROJECT_ID = 'trip-cost-f402c';
var TRIP_ID = 'japan-2025-12';
var PEOPLE = ['小鄧', '彥甫', '煒琮', '皇奇'];
var ID_COL_HEADER = 'FirestoreID'; // 會自動加在試算表最後一欄，用來記錄「這列已經同步過」
// ======================================

function doGet(e) {
  var result = syncAll();
  return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.TEXT);
}

function firestoreBaseUrl() {
  return 'https://firestore.googleapis.com/v1/projects/' + FIREBASE_PROJECT_ID +
    '/databases/(default)/documents/trips/' + TRIP_ID + '/expenses';
}

function syncAll() {
  var toFirestore = syncSheetToFirestore();
  var toSheet = syncFirestoreToSheet();
  return '同步完成\n從 Sheet 新增到資料庫：' + toFirestore + ' 筆\n從網頁新增到 Sheet：' + toSheet + ' 筆';
}

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

// 找到（或建立）用來記錄同步狀態的欄位，回傳 1-based 欄號
function ensureIdColumn(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = headers.indexOf(ID_COL_HEADER);
  if (idx === -1) {
    idx = headers.length;
    sheet.getRange(1, idx + 1).setValue(ID_COL_HEADER);
  }
  return idx + 1;
}

function getHeaderMap(sheet, idCol) {
  var headers = sheet.getRange(1, 1, 1, idCol).getValues()[0];
  var col = {};
  headers.forEach(function (h, i) { if (h) col[h] = i; });
  return col;
}

function formatDate(val) {
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val || '');
}

// ---- Sheet 裡「還沒有 FirestoreID」的新列 → 寫進 Firestore ----
function syncSheetToFirestore() {
  var sheet = getSheet();
  var idCol = ensureIdColumn(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  var col = getHeaderMap(sheet, idCol);
  var data = sheet.getRange(2, 1, lastRow - 1, idCol).getValues();
  var count = 0;

  for (var r = 0; r < data.length; r++) {
    var row = data[r];
    var existingId = row[idCol - 1];
    var item = col['項目'] !== undefined ? row[col['項目']] : '';
    var amount = col['金額（台幣）'] !== undefined ? row[col['金額（台幣）']] : '';
    if (existingId || !item || !amount) continue;

    var participants = PEOPLE.filter(function (name) {
      return col[name] !== undefined && row[col[name]] === true;
    });

    var payload = {
      fields: {
        type: { stringValue: 'shared' },
        source: { stringValue: 'sheet' },
        date: { stringValue: formatDate(row[col['日期']]) },
        item: { stringValue: String(item) },
        currency: { stringValue: 'TWD' },
        amount: { doubleValue: Number(amount) },
        amountTWD: { doubleValue: Number(amount) },
        payer: { stringValue: String(row[col['代墊人']] || '') },
        participants: {
          arrayValue: { values: participants.map(function (n) { return { stringValue: n }; }) }
        },
        method: { stringValue: String(row[col['付款方式']] || '') },
        createdAt: { timestampValue: new Date().toISOString() }
      }
    };

    var resp = UrlFetchApp.fetch(firestoreBaseUrl(), {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var json = JSON.parse(resp.getContentText());
    if (json.name) {
      var newId = json.name.split('/').pop();
      sheet.getRange(r + 2, idCol).setValue(newId);
      count++;
    }
  }
  return count;
}

// ---- 從網頁新增、且還沒同步過的共同花費 → 新增一列到 Sheet ----
function listWebExpenses() {
  var url = firestoreBaseUrl() + '?pageSize=200';
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var json = JSON.parse(resp.getContentText());
  var docs = json.documents || [];
  return docs.filter(function (doc) {
    var f = doc.fields || {};
    var type = f.type ? f.type.stringValue : '';
    var source = f.source ? f.source.stringValue : 'web';
    var sheetSynced = f.sheetSynced ? f.sheetSynced.booleanValue : false;
    return type === 'shared' && source !== 'sheet' && !sheetSynced;
  });
}

function markSynced(docId) {
  var url = firestoreBaseUrl() + '/' + docId + '?updateMask.fieldPaths=sheetSynced';
  var payload = { fields: { sheetSynced: { booleanValue: true } } };
  UrlFetchApp.fetch(url, {
    method: 'patch',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

function numberField(f) {
  if (!f) return 0;
  if (typeof f.doubleValue === 'number') return f.doubleValue;
  if (typeof f.integerValue !== 'undefined') return Number(f.integerValue);
  return 0;
}

function syncFirestoreToSheet() {
  var sheet = getSheet();
  var idCol = ensureIdColumn(sheet);
  var col = getHeaderMap(sheet, idCol);
  var docs = listWebExpenses();
  var count = 0;

  docs.forEach(function (doc) {
    var f = doc.fields || {};
    var docId = doc.name.split('/').pop();

    var currency = f.currency ? f.currency.stringValue : 'TWD';
    var amountTWD = numberField(f.amountTWD);
    var itemLabel = f.item ? f.item.stringValue : '';
    if (currency === 'JPY') {
      itemLabel += '（¥' + numberField(f.amount) + ' 換算）';
    }
    var participants = ((f.participants && f.participants.arrayValue.values) || [])
      .map(function (v) { return v.stringValue; });

    var lastCol = Math.max(sheet.getLastColumn(), idCol);
    var fullRow = new Array(lastCol).fill('');
    if (col['日期'] !== undefined) fullRow[col['日期']] = f.date ? f.date.stringValue : '';
    if (col['項目'] !== undefined) fullRow[col['項目']] = itemLabel;
    if (col['金額（台幣）'] !== undefined) fullRow[col['金額（台幣）']] = amountTWD;
    if (col['代墊人'] !== undefined) fullRow[col['代墊人']] = f.payer ? f.payer.stringValue : '';
    if (col['付款方式'] !== undefined) fullRow[col['付款方式']] = f.method ? f.method.stringValue : '';
    if (col['分攤人數'] !== undefined) fullRow[col['分攤人數']] = participants.length;
    PEOPLE.forEach(function (name) {
      if (col[name] !== undefined) fullRow[col[name]] = participants.indexOf(name) !== -1;
    });
    if (col['每人分攤金額'] !== undefined) {
      fullRow[col['每人分攤金額']] = participants.length ? Math.floor(amountTWD / participants.length) : 0;
    }
    fullRow[idCol - 1] = docId;

    var newRowIndex = sheet.getLastRow() + 1;
    sheet.getRange(newRowIndex, 1, 1, fullRow.length).setValues([fullRow]);

    markSynced(docId);
    count++;
  });

  return count;
}
