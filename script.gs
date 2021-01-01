// ====================================
// get object
// ====================================
var objSheet = SpreadsheetApp.getActiveSpreadsheet();

// ====================================
// get active sheet
// ====================================
var sheet = objSheet.getActiveSheet();

// ====================================
// get active cell
// ====================================
var rowNum = objSheet.getActiveCell().getRow();
var colNum = objSheet.getActiveCell().getColumn();
var valStr = objSheet.getActiveCell().getValue();

// ====================================
// const result word
// ====================================
var complete = "完了";

// ====================================
// const header name
// ====================================
var colTask = 2;      // 項目
var colStartTime = 3; // 開始
var colResult = 4;    // 結果
var colEndTime = 5;   // 終了
var colWorkTime = 7;  // ワークタイム 



// ====================================
// check active cell
// ====================================
function isTaskColumn(col) {
  return (col === colTask ? true : false);
}

function isComplete(val, col){
  return (val === complete && col === colResult ? true : false);
}


// ====================================
// cell constructor
// ====================================
function CellData(col, row) {
  return {
    col: col,
    row: row,
    setTime: function(colEndTime) {
      sheet.getRange(this.row, colEndTime).setValue(new Date());
    },
    getValue: function() {
      return sheet.getRange(this.row, this.col).getValue();
    }
  }
}


// ====================================
// main
// ====================================
var Cell = new CellData(colNum, rowNum);

if (isComplete(Cell.getValue(), Cell.col)) {
  Browser.msgBox("結果入力");
  Cell.setTime(colEndTime);
}



