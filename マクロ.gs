/** @OnlyCurrentDoc */

/**
 * セルの値を変更したときに呼ばれる。
 * スクリプトエディタから実行しても動作しない
 * */
function cellValueChanged() {
/*
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getSheetByName("todo").activate();
  var activatedCell = spreadsheet.getActiveCell();

  // alter color by status
  var activeColumnIndex = activatedCell.getColumn();
  console.info("active column index: " + activeColumnIndex.toString());

  if (activeColumnIndex == 7) {
    var rowIndex = spreadsheet.getCurrentCell().getRowIndex();
    setRowBackgroundByStatus(rowIndex);
  }
*/
  var spreadsheet = SpreadsheetApp.getActive();
  var todoSheet = spreadsheet.getSheetByName("todo");
  var activatedCell = todoSheet.getActiveCell();

  // alter color by status
  var activeColumnIndex = activatedCell.getColumn();
  console.info("active column index: " + activeColumnIndex.toString());
  
  if (activeColumnIndex == 7) {
    var rowIndex = spreadsheet.getCurrentCell().getRowIndex();
    console.info("current row index: " + rowIndex.toString());
    setRowBackgroundByStatus(rowIndex);
  }

}

/** 行の背景色を変更する。色はステータスで決定する。 */
function setRowBackgroundByStatus(rowIndex) {
  if (rowIndex == null) // test用
    rowIndex = 5;

  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getSheetByName("todo").activate();
  var activeSheet = spreadsheet.getActiveSheet();

  var statusColumnIndex = spreadsheet.getRangeByName("status").getColumn();
  console.info("column index at status: " + statusColumnIndex.toString());

  // 選択セルを保存
  //var selectedCell = spreadsheet.getCurrentCell();

  console.info("target　row index: " + rowIndex.toString());
  spreadsheet.getRange('A' + rowIndex.toString() + ':R' + rowIndex.toString()).activate();

  var hexColor = getHexColor(activeSheet.getRange(rowIndex, statusColumnIndex).getValue());
  console.log("targetColor:" + hexColor);
  spreadsheet.getActiveRangeList().setBackground(hexColor);

  // 元に戻す
  //selectedCell.activate();
};

/** ステータスごとの色を取得する */
function getHexColor(status) {
  console.info("status as text: " + status);

  switch (status) {
    case "未対応":
      return "#FF6D01";   // オレンジ
    case "処理中":
      return "#FFFF00";   // 黄
    case '対応検討':
      return "#9900FF";   // 紫
    case "保留":
      return "#00FFFF";   // アクア
    case "処理済み":
      return "#00FF00";   // 緑
    case "完了":
      return "#999999";   // 灰色
    
    case "本日":
      return "#EA4335";   // 赤（少し薄い）
    case "未到来":
      return "#FF00FF";   // 紫（少し赤い）
    default:
      return "#FFFFFF";   // 白
  }
}

/** 新しいタスクを追加する */
function addNewItem() {
  var sheet = SpreadsheetApp.getActive().getActiveSheet();
  var activeRowIndex = sheet.getActiveCell().getRowIndex();
  console.info("activeRow index: " + activeRowIndex.toString());

  // 選択行の下に新タスクを追加
  sheet.insertRowsAfter(activeRowIndex, 1);
  var targetRowIndex = activeRowIndex + 1;
  
  sheet.getRange(targetRowIndex, 1).setValue("=ROW()-1");
  sheet.getRange(targetRowIndex, 6).setValue("橋本");
  sheet.getRange(targetRowIndex, 7).setValue("未対応");

  // 色を変える
  setRowBackgroundByStatus(targetRowIndex);
};