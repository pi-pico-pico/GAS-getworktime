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

function isResultColumn(col) {
  return (col === colResult ? true : false);
}

function isComplete(val){
  return (val === complete ? true : false);
}


// ====================================
// write data
// ====================================
function setWorkTime(col, message) {

  var time = Cell.getColValue(col);
  if (!time) {
    Browser.msgBox(message);
    Cell.setTime(col);
  }
}


// ====================================
// cell constructor
// ====================================
function CellData(col, row) {
  return {
    col: col,
    row: row,
    setTime: function(column) {
      sheet.getRange(this.row, column).setValue(new Date());
    },
    getColValue: function(column) {
      return sheet.getRange(this.row, column).getValue();
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
if (isTaskColumn(Cell.col)) {
  setWorkTime(colStartTime, "開始時間入力");

} else if (isComplete(Cell.getValue()) && isResultColumn(Cell.col)) {
  setWorkTime(colEndTime, "終了時間入力");

}



