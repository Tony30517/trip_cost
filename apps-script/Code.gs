// ==== 設定：換其他行程時只要改這裡 ====
var FIREBASE_PROJECT_ID = 'trip-cost-f402c';
var TRIP_ID = 'japan-2025-12';
var PEOPLE = ['小鄧', '彥甫', '煒琮', '皇奇'];
// ======================================
//
// 同步邏輯（手動觸發，按一次跑一次）：
// 1. 把網頁新增、Sheet 裡還沒有的「共同花費」寫成新的一列到 Sheet
// 2. 把資料庫裡所有「共同花費」全部刪掉
// 3. 完全照 Sheet 目前的每一列，重新建立一次「共同花費」
//
// 這樣 Sheet 上的刪除、改金額，下次同步時會自動反映到資料庫，
// 不用另外追蹤「哪一列同步過」。個人花費不受影響（Sheet 沒有這個概念）。

function doGet(e) {
  var result = syncAll();
  return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.TEXT);
}

function firestoreBaseUrl() {
  return 'https://firestore.googleapis.com/v1/projects/' + FIREBASE_PROJECT_ID +
    '/databases/(default)/documents/trips/' + TRIP_ID + '/expenses';
}

function syncAll() {
  var pushed = pushWebExpensesToSheet();
  var rebuilt = rebuildSharedExpensesFromSheet();
  return '同步完成\n網頁新增寫入 Sheet：' + pushed + ' 筆\n依 Sheet 內容重建共同花費：' + rebuilt + ' 筆';
}

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

function getHeaderMap(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
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

function numberField(f) {
  if (!f) return 0;
  if (typeof f.doubleValue === 'number') return f.doubleValue;
  if (typeof f.integerValue !== 'undefined') return Number(f.integerValue);
  return 0;
}

function listAllExpenses() {
  var url = firestoreBaseUrl() + '?pageSize=300';
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var json = JSON.parse(resp.getContentText());
  return json.documents || [];
}

// ---- 步驟 1：網頁新增（source = web）的共同花費 → 寫成 Sheet 新的一列 ----
function pushWebExpensesToSheet() {
  var sheet = getSheet();
  var col = getHeaderMap(sheet);

  var webDocs = listAllExpenses().filter(function (doc) {
    var f = doc.fields || {};
    var type = f.type ? f.type.stringValue : '';
    var source = f.source ? f.source.stringValue : 'web';
    return type === 'shared' && source === 'web';
  });

  webDocs.forEach(function (doc) {
    var f = doc.fields || {};
    var currency = f.currency ? f.currency.stringValue : 'TWD';
    var amountTWD = numberField(f.amountTWD);
    var itemLabel = f.item ? f.item.stringValue : '';
    if (currency === 'JPY') {
      itemLabel += '（¥' + numberField(f.amount) + ' 換算）';
    }
    var participants = ((f.participants && f.participants.arrayValue.values) || [])
      .map(function (v) { return v.stringValue; });

    var lastCol = Math.max(sheet.getLastColumn(), 1);
    var row = new Array(lastCol).fill('');
    if (col['日期'] !== undefined) row[col['日期']] = f.date ? f.date.stringValue : '';
    if (col['項目'] !== undefined) row[col['項目']] = itemLabel;
    if (col['幣種'] !== undefined) row[col['幣種']] = '台幣';
    if (col['金額'] !== undefined) row[col['金額']] = amountTWD;
    if (col['代墊人'] !== undefined) row[col['代墊人']] = f.payer ? f.payer.stringValue : '';
    if (col['付款方式'] !== undefined) row[col['付款方式']] = f.method ? f.method.stringValue : '';
    if (col['分攤人數'] !== undefined) row[col['分攤人數']] = participants.length;
    PEOPLE.forEach(function (name) {
      if (col[name] !== undefined) row[col[name]] = participants.indexOf(name) !== -1;
    });
    if (col['每人分攤金額'] !== undefined) {
      row[col['每人分攤金額']] = participants.length ? Math.floor(amountTWD / participants.length) : 0;
    }

    sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  });

  return webDocs.length;
}

// ---- 步驟 2+3：刪掉資料庫裡所有共同花費，照 Sheet 目前內容重建 ----
function rebuildSharedExpensesFromSheet() {
  var existing = listAllExpenses().filter(function (doc) {
    var f = doc.fields || {};
    return (f.type ? f.type.stringValue : '') === 'shared';
  });
  existing.forEach(function (doc) {
    UrlFetchApp.fetch(
      'https://firestore.googleapis.com/v1/' + doc.name,
      { method: 'delete', muteHttpExceptions: true }
    );
  });

  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  var col = getHeaderMap(sheet);
  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var count = 0;

  for (var r = 0; r < data.length; r++) {
    var row = data[r];
    var item = col['項目'] !== undefined ? row[col['項目']] : '';
    var amount = col['金額'] !== undefined ? row[col['金額']] : '';
    if (!item || !amount) continue;

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
    if (JSON.parse(resp.getContentText()).name) count++;
  }
  return count;
}
