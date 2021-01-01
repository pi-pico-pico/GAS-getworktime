// ====================================
// シートオブジェクトの取得
// ====================================
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

// ====================================
// アクティブセル（列）の取得
// ====================================
var colNum = spreadsheet.getActiveCell().getColumn();

// ====================================
// アクティブセル（値）の取得
// ====================================
var value = spreadsheet.getActiveCell().getValue();

// ====================================
// アルファベット列名
// ====================================
var colNameB = 2;
var colNameD = 4;

// ====================================
// 結果
// ====================================
var complet = "完了";


// ====================================
// アクティブセル（B列・D列）の確認
// ====================================
function isTargetColumn(col) {
  return (col === colNameB ? true
        : col === colNameD ? true
        : false);
}

// ====================================
// 項目の結果確認
// ====================================
function isComplet(val){
  return (val === complet ? true : false);
}

Browser.msgBox(isComplet(value));