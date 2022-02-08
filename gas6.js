const UPDATECOL = 3;
const STRINGCOL = 4;
const MEMBERCOL = 6;
const COMPLETEDATECOL = 7;
const LIMITDATECOL =8;
const STATUSCOL = 9;
const CHANGEDATECOL = 12;
const DELETECOL = 13;
const STATUS = ['未対応', '対応中', '対応済み', '完了'];

// function showDialog() {
//   　var output = HtmlService.createTemplateFromFile('index');
//   　var ss = SpreadsheetApp.getActiveSpreadsheet();
//   　var html = output.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
//              .setWidth(500)
//              .setHeight(300)
//              .setTitle('モーダルウィンドウテスト');
//   　ss.show(html);    //メッセージボックスとしてを表示する
// }


function getSheet(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  return sheet;
}

function getTaskCell() {
  const sheet = getSheet('プロジェクト_中村');
  const lastRow = sheet.getLastRow();

  for (let i = 7;i <= lastRow; i++) {
    const value = sheet.getRange(i, stringCol).getValue();
    if (value != '') {
      insertMemberList(i, MEMBERCOL, sheet);
      insertStatusList(i, STATUSCOL, sheet);
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

  for (let i = 1;i <= lastRow; i++) {
    const updateRange = sheet.getRange(i, UPDATECOL);
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
  const value = sheet.getRange(row, STATUSCOL).getValue();
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

function deleteTask(deleteCells, sheet) {
  deleteCells.forEach(row => {
      sheet.deleteRows(row);
  })
}

function changeDate(changeDateRow, sheet) {
  const now = new Date();
  const month = now.getMonth() + 1;
  const date = now.getDate();
  sheet.getRange(changeDateRow, CHANGEDATECOL).setValue(month + '/' + date);
}

function isntBlank(val) {
  if (val[0] != '') return val
}

function showLimit() {
  const sheet = getSheet('プロジェクト_中村');
  // const sheet = SpreadsheetApp.getActiveSheet()
  const val = sheet.getRange(1,COMPLETEDATECOL, sheet.getLastRow()).getValues();
  const values = val.filter(isntBlank);
  for (let i = 0; i < values.length; i++) {
    console.log(values[i].getFullYear())
    // const target = new Date(values.getFullYear(), Val.getMonth(), Val.getDate())
    const now = new Date()
    console.log(now.getFullYear())
    // const today = new Date(now.getFullYear(),now.getMonth(),now.getDate())
    // console.log((target - today)/60*60*24)
    // return (target - today)/60*60*24
  }
}
