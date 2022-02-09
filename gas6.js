const IS_UPDATE_COL = 2;
const DATA_START = 3;
const STATUS_COL = 8;
const IS_DELETE_COL = 12;
const STATUS = {unsupported: "未対応", progress: "対応中", complete: "対応済み", finish: "完了"};
const TASK_SHEET = 'プロジェクト_中村';
const colors = {white: "#FFFFFF", green: "#d9ead3", blue: "#cfe2f3", gray: "#d9d9d9", red: "#F08080", yellow: "#FFFACD"};

function getDiffDate(i){
  const DUE_COL =7;
  let sheet = getSheet(TASK_SHEET);
  let today = new Date();
  let date2 = new Date(sheet.getRange(i, DUE_COL).getValue());

  let diffDate = (date2 - today) / (60 * 60 * 24 * 1000);
  return diffDate;
}

function setMark(){
  let sheet = getSheet(TASK_SHEET);
  let lastRow = sheet.getLastRow();
  let warning = [];
  let attention = [];
  
  for(let i= DATA_START; i <= lastRow; i++){
    let diffDate = getDiffDate(i);
    let status = sheet.getRange(i, STATUS_COL).getDisplayValue();
    let compStatus = STATUS.finish;
    let taskName = sheet.getRange(i, DATA_START).getValue();
    if(diffDate < 1 && status != compStatus){
      sheet.getRange(i, DATA_START).setBackground(colors.red);
      warning.push(taskName);
    }else if(diffDate < 3 && status != compStatus){
      sheet.getRange(i, DATA_START).setBackground(colors.yellow);
      attention.push(taskName);
    }else if(status == compStatus) {
      sheet.getRange(i, DATA_START).setBackground(colors.gray);      
    }  
  }
  return [warning, attention];
}

function createSlackMessage() {
  notifications = setMark();
  warningTasks = notifications[0].join("、");
  attentionTasks = notifications[1].join("、");
  message = "締め切りが過ぎているタスク：　" + warningTasks + "\n 締め切り間近なタスク：　" + attentionTasks; 
  
  return message;
}

function getSheet(sheetName) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  return sheet;
}

function createMemberList() {
  const USER_SHEET = 'メンバー';
  let memberSheet = getSheet(USER_SHEET);
  let lastRow = memberSheet.getLastRow();
  let startCol = 1;
  let startRow = 2;
  let memberRange = memberSheet.getRange(startRow, startCol, lastRow);
  let memberList = SpreadsheetApp.newDataValidation().requireValueInRange(memberRange).build();
  return memberList;
}

function insertMemberList(row, col, sheet) {
  let memberList = createMemberList(); 
  let setCell = sheet.getRange(row,col);
  setCell.setDataValidation(memberList);
}

function getDefaultStatus() {
  defaultStatus = [];
  Object.values(STATUS).forEach(value => {
	defaultStatus.push(value);
  });
  return defaultStatus;
}

function insertStatusList(row, col, sheet) {
  let defaultStatus = getDefaultStatus();
  let statusList = SpreadsheetApp.newDataValidation().requireValueInList(defaultStatus).build();
  let setCell = sheet.getRange(row, col);
  setCell.setDataValidation(statusList);
}

function insertLists(row, sheet) {
  const MEMBER_COL = 5;
  insertMemberList(row, MEMBER_COL, sheet);
  insertStatusList(row, STATUS_COL, sheet);
}


function setDate(row, col, sheet) {
  const now = new Date();
  const month = now.getMonth() + 1;
  const date = now.getDate();
  sheet.getRange(row, col).setValue(month + '/' + date);
}

function getDeleteRowList(lastRow, sheet) {
  let deleteRowList = [];
  for (let i = 3; i <= lastRow; i++) {
    let isDeleteValue = sheet.getRange(i, IS_DELETE_COL).getValue();
    if (isDeleteValue == true) {
      deleteRowList.unshift(i);
    } 
  }
  return deleteRowList;
}

function deleteTask(lastRow, sheet) {
  let deleteRowList = getDeleteRowList(lastRow, sheet);
  deleteRowList.forEach(row => {
    sheet.deleteRow(row);
  });
}

function setStatusColor(row, sheet) {
  let status = sheet.getRange(row, STATUS_COL).getValue();
  let lastCol = sheet.getLastColumn();
  let setColorRange = sheet.getRange(row, DATA_START, 1, lastCol);
    
  if (status == STATUS.unsupported) {
    setColorRange.setBackground(colors.white);
  } else if (status == STATUS.progress) {
    setColorRange.setBackground(colors.green);
  } else if (status == STATUS.complete) {
    setColorRange.setBackground(colors.blue); 
  } else if (status == STATUS.finish) {
    setColorRange.setBackground(colors.gray); 
  }
}

function updateTask(row, sheet){
  const CHANGE_DATE_COL = 11;
  for (let i = 1; i <= row; i++) {
    let updateRange = sheet.getRange(i, IS_UPDATE_COL);
    let isUpdateValue = updateRange.getValue();
    if (isUpdateValue == true) {
      updateRange.uncheck();
      setStatusColor(i, sheet);
      setDate(i, CHANGE_DATE_COL, sheet);
    }
  }
}

function setCreator(createRow, sheet) {
  const CREATOR_COL = 9;
  let message = '作成者の名前を入力して下さい。';
  let creator = Browser.inputBox(message, Browser.Buttons.OK_CANCEL);
  if(creator == 'cancel'){
    return false;
  }else{
    sheet.getRange(createRow, CREATOR_COL).setValue(creator);
    return true;  
  }
}

function addTask() {
  const COMPLETE_DATE_COL = 6;
  const CREATE_DATE_COL = 10;
  let sheet = getSheet(TASK_SHEET);
  let createRow = sheet.getLastRow() + 1;
  isCreate = setCreator(createRow, sheet);
  if(isCreate){
    sheet.getRange(createRow, IS_UPDATE_COL).insertCheckboxes();
    sheet.getRange(createRow, IS_DELETE_COL).insertCheckboxes();
    sheet.getRange(createRow, COMPLETE_DATE_COL, 1, 2).setNumberFormat("MM/dd");
    insertLists(createRow, sheet);
    setDate(createRow, CREATE_DATE_COL, sheet);
  }else{
    Browser.msgBox("タスク作成をキャンセルしました。");
  }
}

function updateSheetInfo() {
  let sheet = getSheet(TASK_SHEET); 
  let lastRow = sheet.getLastRow();
  deleteTask(lastRow, sheet);
  updateTask(lastRow, sheet);
  setMark();
}

function onOpen() {
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu("メニュー");
  menu.addItem("更新","updateSheetInfo");
  menu.addItem("新規タスク作成","addTask");
  menu.addToUi();
}

function slack() {
  const webhookUrl = 'https://hooks.slack.com/services/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';
  const username = 'username';  // 通知時に表示されるユーザー名
  const icon = ':hatching_chick:';  // 通知時に表示されるアイコン
  let message = createSlackMessage();  // 投稿メッセージ
  let jsonData = {"username" : username, "icon_emoji": icon, "text" : message};
  let payload = JSON.stringify(jsonData);
  let options = {"method" : "post", "contentType" : "application/json", "payload" : payload};

  UrlFetchApp.fetch(webhookUrl, options);
}