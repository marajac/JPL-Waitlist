var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheets()[0];
var today = new Date();

var TOP_ROW = 2;
var SEARCH_COLUMN = 1;

var staffInitials = sheet.getRange(3,9).getValue();
var bookTitle = sheet.getRange(4,9).getValue();
var addPagerInput = sheet.getRange(5,9).getValue();
var updatePagerInput = sheet.getRange(8, 8).getValue();
var removePagerInput = sheet.getRange(11, 8).getValue();
var inputField = sheet.getRange(3, 9, 3);

function onEdit(e){
  var source = e.source.getActiveSheet().getActiveCell();
  var sourceCell = source.getA1Notation()
  var sourceValue = source.getValue();
  if(sourceValue === "")
    return;
  else if(sourceCell == 'H8'){
    if(!isNaN(sourceValue)){
      updateItem(updatePagerInput);
    }
  }
  else if(sourceCell == 'H11'){
    if(!isNaN(sourceValue)){
      clearItem(removePagerInput);
    }
  }
  else if(sourceCell == 'I5'){
    if(!isNaN(sourceValue)){
      if(bookTitle == "" || staffInitials == ""){
        Browser.msgBox("BOOK TITLE and STAFF INITIALS are required before adding an item to the wait list.");
        sheet.getRange(5,9).setValue("");
      }
      else{
        addAccept();
      }
    }
    else if(sourceValue.toLowerCase() == "declined"){
      addDecline();
    }
  }
}

function addDecline() {
  var firstEmptyRow = getFirstEmptyRow();
  addPager(firstEmptyRow, 'declined');
  addInfoAndDate(firstEmptyRow);
  formatRow(firstEmptyRow, 'line-through');
  formatInputField();
}

function addAccept() {
  var firstEmptyRow = getFirstEmptyRow();
  sheet.getRange(firstEmptyRow, 2).setBackgroundRGB(255, 255, 0);
  addPager(firstEmptyRow, addPagerInput);
  addInfoAndDate(firstEmptyRow);
  formatRow(firstEmptyRow, 'none');
  formatInputField();
}

function addPager(row, value){
  sheet.getRange(row, 1).setValue(value);
}

function addInfoAndDate(row){
  sheet.getRange(row, 2).setValue(bookTitle);
  sheet.getRange(row, 6).setValue(staffInitials);
  sheet.getRange(row, 3).setValue(today);
  sheet.getRange(row, 4).setValue(today);
  clearItemForm();
}

function formatRow(row, format){
  for(i = 1; i <= 6; i++){
    sheet.getRange(row, i).setFontLine(format);
  }
}

function formatInputField(){
  inputField.setBorder(true, true, true, true, false, true, null, SpreadsheetApp.BorderStyle.SOLID);
  inputField.setBackgroundRGB(255, 255, 255);
  inputField.setFontColor('black');
  inputField.setFontFamily("Arial");
  inputField.setFontLine('none');
  inputField.setFontWeight('normal');
  inputField.setFontStyle('normal');
}

function clearItemForm(){
  sheet.getRange(3,9).setValue("");
  sheet.getRange(4,9).setValue("");
  sheet.getRange(5,9).setValue("");
}

function clearUpdateForm(){
  sheet.getRange(8,8).setValue("");
}

function clearRemoveForm(){
  sheet.getRange(11,8).setValue("");
}

function clearItem(pager){
  var row = getActivePager(pager);
  sheet.getRange(row, 2).setBackgroundRGB(255, 255, 255);
  sheet.getRange(row, 5).setBackgroundRGB(255, 255, 255);
  formatRow(row, 'line-through');
  clearRemoveForm();
}

function updateItem(pager){
  var pickUpTime = today;
  var row = getActivePager(pager);
  pickUpTime.setMinutes(pickUpTime.getMinutes() + 15);
  sheet.getRange(row, 5).setValue(pickUpTime);
  sheet.getRange(row, 5).setBackgroundRGB(255, 255, 0); 
  clearUpdateForm();
}

function getActivePager(pager){
  var firstEmptyRow = getFirstEmptyRow();
  for(i = TOP_ROW; i < firstEmptyRow; i++){
      if(isActivePager(i, pager))
        return i;
  }
  Browser.msgBox("Pager " + pager + " not found, please check your entry and try again.");
}

function isActivePager(row, pager){
  return (sheet.getRange(row, SEARCH_COLUMN).getValue() == pager &&
          sheet.getRange(row, SEARCH_COLUMN).getFontLine() != 'line-through');
}

function getFirstEmptyRow() {
  var column = ss.getRange('A:A');
  var values = column.getValues(); // get all data in one call
  var ct = TOP_ROW;
  while (sheet.getRange(ct, SEARCH_COLUMN).getValue() != "") {
    ct++;
  }
  return (ct);
}
