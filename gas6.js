const UPDATE_COL = 3;
const STRING_COL = 4;
const MEMBER_COL = 6;
const COMPLETE_DATE_COL = 7;
const LIMIT_DATE_COL =8;
const STATUS_COL = 9;
const CHANGE_DATE_COL = 12;
const DELETE_COL = 13;
const STATUS = ['未対応', '対応中', '対応済み', '完了'];
// slack連携
const webhookUrl = 'https://hooks.slack.com/services/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';
const username = 'username';  // 通知時に表示されるユーザー名
const icon = ':hatching_chick:';  // 通知時に表示されるアイコン
let message = 'test';  // 投稿メッセージ

function getSheet(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().  getSheetByName(sheetName);
  return sheet;
}

function getTaskCell() {
  const sheet = getSheet('プロジェクト_中村');
  const lastRow = sheet.getLastRow();

  for (let i = 7;i <= lastRow; i++) {
    const value = sheet.getRange(i, stringCol).getValue();
    if (value != '') {
      insertMemberList(i, MEMBER_COL, sheet);
      insertStatusList(i, STATUS_COL, sheet);
    }
  }
}

function createMemberList() {
  const memberSheet = getSheet('メンバー');
  const lastRow = memberSheet.getLastRow();
  const memberSheetCol = 2;
  const memberSheetRow = 2;
  const memberRange = memberSheet.getRange(memberSheetRow, memberSheetCol, lastRow);
  const memberList = SpreadsheetApp.newDataValidation().requireValueInRange(memberRange).build();
  return memberList;
}

function insertMemberList(row, col, sheet) {
  const memberList = createMemberList(); 
  const cell = sheet.getRange(row,col);
  cell.setDataValidation(memberList);
}
　
function insertStatusList(row, col, sheet) {
  const statusList = SpreadsheetApp.newDataValidation().requireValueInList(STATUS).build();
  const cell = sheet.getRange(row, col);
  cell.setDataValidation(statusList);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu("メニュー");
  menu.addItem("実行","reflectChange");
  menu.addToUi();
}

function reflectChange() {
  const sheet = getSheet('プロジェクト_中村');
  const lastRow = sheet.getLastRow();
  const deleteCells = [];

  for (let i = 1; i <= lastRow; i++) {
    const updateRange = sheet.getRange(i, UPDATE_COL);
    const deleteRange = sheet.getRange(i, DALETECOL);
    
    if (deleteRange.getValue() == true) {
      deleteCells.unshift(i);
    } else if (updateRange.getValue() == true) {
      updateRange.uncheck();
      changeStatusColor(i, sheet);
      changeDate(i, sheet);
    }
  }
  deleteTask(deleteCells, sheet)
}

function changeStatusColor(row, sheet) {
  const value = sheet.getRange(row, STATUS_COL).getValue();
  const lastCol = sheet.getLastColumn();

  if (value == STATUS[0]) {
    sheet.getRange(row, STRINGCOL, 1, lastCol).setBackground("#FFFFFF");
  } else if (value == STATUS[1]) {
    sheet.getRange(row, STRINGCOL, 1, lastCol).setBackground("#d9ead3");
  } else if (value == STATUS[2]) {
    sheet.getRange(row, STRINGCOL, 1, lastCol).setBackground("#cfe2f3"); 
  } else if (value == STATUS[3]) {
    sheet.getRange(row, STRINGCOL, 1, lastCol).setBackground("#d9d9d9"); 
  }
}

function deleteTask( deleteCells, sheet ) {
  deleteCells.forEach(row => {
      sheet.deleteRows(row);
  })
}

function changeDate(changeDateRow, sheet) {
  const now = new Date();
  const month = now.getMonth() + 1;
  const date = now.getDate();
  sheet.getRange( changeDateRow, CHANGE_DATE_COL ).setValue( month + '/' + date_ );
}

function isntBlank(val) {
  if (val[0] != '') return val
}

function showLimit() {
  const sheet = getSheet('プロジェクト_中村');
  // const sheet = SpreadsheetApp.getActiveSheet()
  const val = sheet.getRange( 1, COMPLETE_DATE_COL, sheet.getLastRow() ).ge_tVal_ues();
  const values = val.filter(isntBlank);
  for (let i = 0; i < values.length; i++) {
    // const target = new Date(values.getFullYear(), Val.getMonth(), Val.getDate())
    const now = new Date()
    // const today = new Date(now.getFullYear(),now.getMonth(),now.getDate())
    // console.log((target - today)/60*60*24)
    // return (target - today)/60*60*24
  }
}

function changeColor( row, col , color) {
  let sheet = getSheet('プロジェクト_中村')
  sheet.getRange(row, col).setFontColor( color )
}

function getCellValue( row, col) {
  let sheet = getSheet('プロジェクト_中村')
  let cellValue = sheet.getRange(row, col).getValue()
  return cellValue
}

function addString(row, col) {
  let pointedValue = getCellValue(row, col) +  '●'
  return pointedValue
}

// slack連携

function slack() {
  let jsonData =
  {
     "username" : username,
     "icon_emoji": icon,
     "text" : message
  };
  let payload = JSON.stringify(jsonData);

  let options =
  {
    "method" : "post",
    "contentType" : "application/json",
    "payload" : payload
  };

  UrlFetchApp.fetch(webhookUrl, options);
}
