const IS_UPDATE_COL = 2;
const DATA_START = 3;
const MEMBER_COL = 5;
const COMPLETE_DATE_COL = 6;
const DUE_COL =7;
const STATUS_COL = 8;
const CREATER_COL = 9;
const CREATE_DATE_COL = 10;
const CHANGE_DATE_COL = 11;
const IS_DELETE_COL = 12;
const STATUS = ['未対応', '対応中', '対応済み', '完了'];
const TASK_SHEET = 'プロジェクト_中村';
const USER_SHEET = 'メンバー';

function test() {
  date = Utilities.formatDate(new Date(), "Asia/Tokyo", "MM/dd");
  console.log(new Date());
}

function isString(value) {
  if (typeof value === "string" || value instanceof String) {
    return true;
  } else {
    return false;
  }
}

function getDateByString(_date){
  if(isString(_date)){
    _date = _date.replace('⚫︎ ', '');
  } else {
    _date = Utilities.formatDate(new Date(_date), "Asia/Tokyo", "MM/dd");
  }
  return _date;
}

function getDiffDate(i){
  let sheet = getSheet(TASK_SHEET);
  let today = new Date(Utilities.formatDate(new Date(), "Asia/Tokyo", "MM/dd"));
  let date2 = new Date(getDateByString(sheet.getRange(i, DUE_COL).getValue()));

  let diffDate = (date2 - today) / (60 * 60 * 24 * 1000);
  return diffDate;
}

function setMark(){
  let sheet = getSheet(TASK_SHEET);
  lastRow = sheet.getLastRow();
  for(let i= DATA_START; i <= lastRow; i++){
    let diffDate = getDiffDate(i);
    let status = sheet.getRange(i, STATUS_COL).getDisplayValue();
    let compStatus = STATUS[3];
    if(diffDate < 1 && status != compStatus){
      sheet.getRange(i, DATA_START).setBackground("#FF3333");
    }else if(diffDate < 3 && status != compStatus){
      sheet.getRange(i, DATA_START).setBackground("#FFFF66");
    }
  }
}

function getSheet(sheetName) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  return sheet;
}

function createMemberList() {
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

function insertStatusList(row, col, sheet) {
  let statusList = SpreadsheetApp.newDataValidation().requireValueInList(STATUS).build();
  let setCell = sheet.getRange(row, col);
  setCell.setDataValidation(statusList);
}

function insertLists(row, sheet) {
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
  for (let i = 1; i <= lastRow; i++) {
    let isDeleteValue = sheet.getRange(i, IS_DELETE_COL).getValue();
    if (isDeleteValue == true) {
      deleteRowList.push(i);
    } 
  }
  return deleteRowList;
}

function deleteTask(lastRow, sheet) {
  let deleteRowList = getDeleteRowList(lastRow, sheet);
  deleteRowList.forEach(row => {
    sheet.deleteRow(row);
  })
}

function setStatusColor(row, sheet) {
  const colors = ["#FFFFFF", "#d9ead3", "#cfe2f3", "#d9d9d9"];
  let status = sheet.getRange(row, STATUS_COL).getValue();
  let lastCol = sheet.getLastColumn();
  let setColorRange = sheet.getRange(row, DATA_START, 1, lastCol);
    
  if (status == STATUS[0]) {
    setColorRange.setBackground(colors[0]);
  } else if (status == STATUS[1]) {
    setColorRange.setBackground(colors[1]);
  } else if (status == STATUS[2]) {
    setColorRange.setBackground(colors[2]); 
  } else if (status == STATUS[3]) {
    setColorRange.setBackground(colors[3]); 
  }
}

function updateTask(row, sheet){
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

function setCreater(createRow, sheet) {
  let message = '作成者の名前を入力して下さい。';
  let creater = Browser.inputBox(message);
  if(creater == cancel){
    return;
  }else{
    sheet.getRange(createRow, CREATER_COL).setValue(creater);
  }
}

function addTask() {
  let sheet = getSheet(TASK_SHEET);
  let createRow = sheet.getLastRow() + 1;
  sheet.getRange(createRow, IS_UPDATE_COL).insertCheckboxes();
  sheet.getRange(createRow, IS_DELETE_COL).insertCheckboxes();
  sheet.getRange(createRow, COMPLETE_DATE_COL, 1, 2).setNumberFormat("MM/dd");
  insertLists(createRow, sheet);
  setDate(createRow, CREATE_DATE_COL, sheet);
  setCreater(createRow, sheet)
}

function updateInfo() {
  let sheet = getSheet(TASK_SHEET);
  let lastRow = sheet.getLastRow();
  deleteTask(lastRow, sheet);
  updateTask(lastRow, sheet);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu("メニュー");
  menu.addItem("更新","updateInfo");
  menu.addItem("新規タスク作成","addTask");
  menu.addToUi();
}
const IS_UPDATE_COL = 2;
const DATA_START = 3;
const MEMBER_COL = 5;
const COMPLETE_DATE_COL = 6;
const DUE_COL =7;
const STATUS_COL = 8;
const CREATER_COL = 9;
const CREATE_DATE_COL = 10;
const CHANGE_DATE_COL = 11;
const IS_DELETE_COL = 12;
const STATUS = ['未対応', '対応中', '対応済み', '完了'];
const TASK_SHEET = 'プロジェクト_中村';
const USER_SHEET = 'メンバー';

function test() {
  date = Utilities.formatDate(new Date(), "Asia/Tokyo", "MM/dd");
  console.log(new Date());
}

function isString(value) {
  if (typeof value === "string" || value instanceof String) {
    return true;
  } else {
    return false;
  }
}

function getDateByString(_date){
  if(isString(_date)){
    _date = _date.replace('⚫︎ ', '');
  } else {
    _date = Utilities.formatDate(new Date(_date), "Asia/Tokyo", "MM/dd");
  }
  return _date;
}

function getDiffDate(i){
  let sheet = getSheet(TASK_SHEET);
  let today = new Date(Utilities.formatDate(new Date(), "Asia/Tokyo", "MM/dd"));
  let date2 = new Date(getDateByString(sheet.getRange(i, DUE_COL).getValue()));

  let diffDate = (date2 - today) / (60 * 60 * 24 * 1000);
  return diffDate;
}

function setMark(){
  let sheet = getSheet(TASK_SHEET);
  lastRow = sheet.getLastRow();
  for(let i= DATA_START; i <= lastRow; i++){
    let diffDate = getDiffDate(i);
    let status = sheet.getRange(i, STATUS_COL).getDisplayValue();
    let compStatus = STATUS[3];
    if(diffDate < 1 && status != compStatus){
      sheet.getRange(i, DATA_START).setBackground("#FF3333");
    }else if(diffDate < 3 && status != compStatus){
      sheet.getRange(i, DATA_START).setBackground("#FFFF66");
    }
  }
}

function getSheet(sheetName) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  return sheet;
}

function createMemberList() {
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

function insertStatusList(row, col, sheet) {
  let statusList = SpreadsheetApp.newDataValidation().requireValueInList(STATUS).build();
  let setCell = sheet.getRange(row, col);
  setCell.setDataValidation(statusList);
}

function insertLists(row, sheet) {
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
  for (let i = 1; i <= lastRow; i++) {
    let isDeleteValue = sheet.getRange(i, IS_DELETE_COL).getValue();
    if (isDeleteValue == true) {
      deleteRowList.push(i);
    } 
  }
  return deleteRowList;
}

function deleteTask(lastRow, sheet) {
  let deleteRowList = getDeleteRowList(lastRow, sheet);
  deleteRowList.forEach(row => {
    sheet.deleteRow(row);
  })
}

function setStatusColor(row, sheet) {
  const colors = ["#FFFFFF", "#d9ead3", "#cfe2f3", "#d9d9d9"];
  let status = sheet.getRange(row, STATUS_COL).getValue();
  let lastCol = sheet.getLastColumn();
  let setColorRange = sheet.getRange(row, DATA_START, 1, lastCol);
    
  if (status == STATUS[0]) {
    setColorRange.setBackground(colors[0]);
  } else if (status == STATUS[1]) {
    setColorRange.setBackground(colors[1]);
  } else if (status == STATUS[2]) {
    setColorRange.setBackground(colors[2]); 
  } else if (status == STATUS[3]) {
    setColorRange.setBackground(colors[3]); 
  }
}

function updateTask(row, sheet){
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

function setCreater(createRow, sheet) {
  let message = '作成者の名前を入力して下さい。';
  let creater = Browser.inputBox(message);
  sheet.getRange(createRow, CREATER_COL).setValue(creater);
}

function addTask() {
  let sheet = getSheet(TASK_SHEET);
  let createRow = sheet.getLastRow() + 1;
  sheet.getRange(createRow, IS_UPDATE_COL).insertCheckboxes();
  sheet.getRange(createRow, IS_DELETE_COL).insertCheckboxes();
  sheet.getRange(createRow, COMPLETE_DATE_COL, 1, 2).setNumberFormat("MM/dd");
  insertLists(createRow, sheet);
  setDate(createRow, CREATE_DATE_COL, sheet);
  setCreater(createRow, sheet)
}

function updateInfo() {
  let sheet = getSheet(TASK_SHEET);
  let lastRow = sheet.getLastRow();
  deleteTask(lastRow, sheet);
  updateTask(lastRow, sheet);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu("メニュー");
  menu.addItem("更新","updateInfo");
  menu.addItem("新規タスク作成","addTask");
  menu.addToUi();
}
