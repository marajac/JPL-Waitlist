/*
 * This file contains all the core code for the waitlist
 * pertaining to its standard operation, i.e.
 * adding/removing items and the logic associated with
 * those actions.
 */

var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheets()[0];
var today = new Date();

var TOP_ROW = 2;
var SEARCH_COLUMN = 1;

var inputField = sheet.getRange(3,9,3);
var updatePagerInput = sheet.getRange(8,8);
var removePagerInput = sheet.getRange(11,8);
var staffInitials = sheet.getRange(3,9).getValue();
var bookTitle = sheet.getRange(4,9).getValue();
var addPagerInput = sheet.getRange(5,9).getValue();

/* TODO: Ensure only the sheet owner has access to this menu */

//function onOpen(){
//  var menuEntries = [{name: "Clean up", functionName: "cleanUpWaitlist"}];
//  ss.addMenu("Wait List", menuEntries);
//}

function onEdit(e){
  var source = e.source.getActiveSheet().getActiveCell();
  var sourceCell = source.getA1Notation()
  var sourceValue = source.getValue();
  if(sourceValue === "")
    return;
  /* Cell I5 is the field where a pager number is entered in order
     to ADD an item to the waitlist */
  else if(sourceCell == 'I5'){
    if(!isNaN(sourceValue)){
      /* Check to see if the Book Title and Staff Initials were entered */
      if(bookTitle == "" || staffInitials == ""){
        Browser.msgBox("BOOK TITLE and STAFF INITIALS are required before adding an item to the wait list.");
        addPagerInput.setValue("");
      }
      else{
        addAccept();
        clearRange(inputField);
      }
    }
    else if(sourceValue.toLowerCase() == "declined"){
      addDecline();
      clearRange(inputField);
    }
  }
  /* Cell H8 is the field where a pager number is entered in order
     to UPDATE an item from the waitlist */
  else if(sourceCell == 'H8'){
    if(!isNaN(sourceValue)){
      updateItem(updatePagerInput.getValue());
      clearRange(updatePagerInput);
    }
  }
  /* Cell H11 is the field where a pager number is entered in order
     to REMOVE an item from the waitlist */
  else if(sourceCell == 'H11'){
    if(!isNaN(sourceValue)){
      clearItem(removePagerInput.getValue());
      clearRange(removePagerInput);
    }
  }
}

/* Add an item to the waitlist that a patron has DECLINED */
function addDecline() {
  var firstEmptyRow = getFirstEmptyRow();
  addPager(firstEmptyRow, 'declined');
  addInfoAndDate(firstEmptyRow);
  setRowFontLine(firstEmptyRow, 'line-through');
}

/* Add an item to the waitlist that a patron has ACCEPTED */
function addAccept() {
  var firstEmptyRow = getFirstEmptyRow();
  sheet.getRange(firstEmptyRow, 2).setBackgroundRGB(255, 255, 0);
  addPager(firstEmptyRow, addPagerInput);
  addInfoAndDate(firstEmptyRow);
  setRowFontLine(firstEmptyRow, 'none');
}

function addPager(row, value){
  sheet.getRange(row, 1).setValue(value);
}

function addInfoAndDate(row){
  sheet.getRange(row, 2).setValue(bookTitle);
  sheet.getRange(row, 6).setValue(staffInitials);
  sheet.getRange(row, 3).setValue(today);
  sheet.getRange(row, 3).setNumberFormat('M/D/YYYY');
  sheet.getRange(row, 4).setValue(today);
  sheet.getRange(row, 4).setNumberFormat('h:mm A/P"M"');
}

function setRowFontLine(row, format){
  sheet.getRange(row, 1,1,6).setFontLine(format);
}

function clearRange(range){
  range.setValue("");
  range.setBorder(true, true, true, true, false, true, null, SpreadsheetApp.BorderStyle.SOLID);
  range.setBackgroundRGB(255, 255, 255);
  range.setFontColor('black');
  range.setFontFamily("Arial");
  range.setFontLine('none');
  range.setFontSize(10);
  range.setFontStyle('normal');
  range.setFontWeight('normal');
}

function clearItem(pager){
  var row = getActivePager(pager);
  sheet.getRange(row, 2).setBackgroundRGB(255,255,255);
  sheet.getRange(row, 5).setBackgroundRGB(255,255,255);
  setRowFontLine(row, 'line-through');
}

function updateItem(pager){
  var pickUpTime = today;
  var row = getActivePager(pager);
  pickUpTime.setMinutes(pickUpTime.getMinutes() + 15);
  sheet.getRange(row, 5).setValue(pickUpTime);
  sheet.getRange(row, 5).setBackgroundRGB(255,255,0);
  sheet.getRange(row, 5).setNumberFormat('h:mm A/P"M"');
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
          sheet.getRange(row, SEARCH_COLUMN).getFontLine() != 'line-through')
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
