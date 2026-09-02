// 在 Sheet 裡的「金額」欄（D欄）輸入日圓金額、按下 Enter 後，
// 自動把該儲存格覆蓋成換算後的台幣金額，並把「幣種」欄（C欄）改回「台幣」。
//
// 事前準備：找一個不礙眼的儲存格（例如 T1）放這條公式，然後把該欄隱藏起來：
//   T1: =GOOGLEFINANCE("CURRENCY:JPYTWD")
// 下面的 RATE_CELL 要跟你實際放的位置一致。

var RATE_CELL = 'T1';
var CURRENCY_COL = 3; // C欄：幣種
var AMOUNT_COL = 4;   // D欄：金額

function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();
  var row = range.getRow();
  var col = range.getColumn();

  // 只處理「金額」欄、單一儲存格的編輯（第2列以後才是資料列）
  if (row < 2 || col !== AMOUNT_COL || range.getNumColumns() > 1 || range.getNumRows() > 1) return;

  var currency = sheet.getRange(row, CURRENCY_COL).getValue();
  if (currency !== '日圓') return;

  var amount = Number(range.getValue());
  if (!amount) return;

  var rate = sheet.getRange(RATE_CELL).getValue();
  if (!rate) return;

  range.setValue(Math.round(amount * rate));
  sheet.getRange(row, CURRENCY_COL).setValue('台幣');
}
