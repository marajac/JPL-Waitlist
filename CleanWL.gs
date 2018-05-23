/*
 * This file contains all the code for cleaning up the waitlist
 * and exporting any removed items to a new spreadsheet for
 * reporting purposes.
 */

/* TODO: Refactor cleanUpWaitlist for readability/efficiency
         Maybe include count per hour */
function cleanUpWaitlist(){
  var firstEmptyRow = getFirstEmptyRow();
  var firstAfterSavedItems, rowToCopyTo;
  var itemsToSave = [];
  var itemsToMove = [];
  var acceptedItems = {};
  var declinedItems = {};
  var rowNum = TOP_ROW;
  var sheetDate = new Date();
  var rangeDate;
  sheetDate.setDate(sheetDate.getDate() - 1);
  /* Iterate through the current list, removing crossed out entries
    and saving those that are still active */
  for(i = TOP_ROW; i <= firstEmptyRow; i++){
    if(sheet.getRange(i, SEARCH_COLUMN).getFontLine() == 'line-through'){
      rangeDate = new Date(sheet.getRange(i, 3).getValue());
      if(rangeDate.getDate() == sheetDate.getDate())
        itemsToMove.push(sheet.getRange(i, 1,1,6));
    }
    else
      itemsToSave.push(sheet.getRange(i, 1,1,6));
  }
  
  /* Count all instances of items that have been removed from the waitlist 
     Accepted items and declined items are counted separately */
  while((i = itemsToMove.shift()) != undefined){
    var pagerValue = i.getValues()[0][0];
    var itemName = i.getValues()[0][1].toLowerCase().trim();
    var nameWithRange = [itemName, i];
    if(pagerValue != 'declined')
      assignKey(acceptedItems, nameWithRange);
    else
      assignKey(declinedItems, nameWithRange);
  }
  
  /* Create a new spreadsheet, containing information from the removed items */
  createReportSheet(acceptedItems, declinedItems);
  
  /* Commenting out this block for testing purposes
  /* Construct a new wait list using the saved active items 
  firstAfterSavedItems = itemsToSave.length + 1;
  while((i = itemsToSave.shift()) != undefined){
    rowToCopyTo = sheet.getRange(rowNum, 1);
    i.copyTo(rowToCopyTo);
    rowNum++;
  }
  
  /* Clean up any rows that still have values after the
     new list is constructed 
  for(i = firstAfterSavedItems; i < firstEmptyRow; i++){
    sheet.getRange(i, 1,1,6).setValue("");
    sheet.getRange(i, 1,1,6).setBackgroundRGB(255,255,255);
  } */
}

function assignKey(obj, key) {
  typeof obj[key[0]] === 'undefined' ? obj[key[0]] = [key[1]] : obj[key[0]].push(key[1]);
}

/* Creates a new spreadsheet formatted and filled with count data
   for items on the waitlist for the previous day */
function createReportSheet(acceptedItems, declinedItems) {
  var row, removedItemsSheet, rowAfterDivider;
  var sheetDate = new Date();
  var sheetDateString = "";
  sheetDate.setDate(sheetDate.getDate() - 1);
  sheetDateString = Utilities.formatDate(sheetDate, "GMT", "M/d/yyyy");
  removedItemsSheet = SpreadsheetApp.create("Waitlist Stats - " + sheetDateString);
  row = TOP_ROW;
  /* Set column size and various formatting for the new sheet */
  removedItemsSheet.setColumnWidth(1, 450);
  removedItemsSheet.setColumnWidth(2, 1);
  removedItemsSheet.setColumnWidth(3, 105);
  removedItemsSheet.setColumnWidth(4, 1);
  removedItemsSheet.setColumnWidth(5, 165);
  removedItemsSheet.getSheets()[0].getRange(1,1,1,5).setFontWeight('bold');
  removedItemsSheet.getSheets()[0].getRange(1,1,1,5).setBackgroundRGB(211, 211, 211);
  removedItemsSheet.getSheets()[0].getRange(1,1,1,5).setHorizontalAlignment('center');
  removedItemsSheet.getRange('A1').setValue("Title/Type");
  removedItemsSheet.getRange('C1').setValue("Times Requested");
  removedItemsSheet.getRange('C1').setWrap(true);
  removedItemsSheet.getRange('E1').setValue("Note");
  
  /* Add stats for the accepted waitlist items */
  for (var key in acceptedItems) {
    removedItemsSheet.getSheets()[0].getRange(row, 1).setValue(key);
    removedItemsSheet.getSheets()[0].getRange(row, 1).setWrap(true);
    removedItemsSheet.getSheets()[0].getRange(row, 3).setValue(acceptedItems[key].length);
    removedItemsSheet.getSheets()[0].getRange(row, 1,1,5).setBorder(true, true, true, true, false, true, null, SpreadsheetApp.BorderStyle.SOLID);
    row++;
  }
  /* Check to see if there are values to sort */
  if(Object.keys(acceptedItems).length > 0)
    removedItemsSheet.getSheets()[0].getRange(2,1,(row-1),5).sort([{column: 3, ascending: false}, {column: 1, ascending: true}]);
  /* Insert a divider between the accepted and declined items */
  removedItemsSheet.getSheets()[0].getRange(row, 1,1,5).setBorder(true, true, true, true, false, true, null, SpreadsheetApp.BorderStyle.SOLID_THICK);
  removedItemsSheet.getSheets()[0].getRange(row, 1,1,5).merge();
  removedItemsSheet.getSheets()[0].getRange(row, 1).setValue("Declined");
  removedItemsSheet.getSheets()[0].getRange(row, 1).setHorizontalAlignment('center');
  removedItemsSheet.getSheets()[0].getRange(row, 1).setBackgroundRGB(255, 165, 0);
  row++;
  
  /* Add stats for the declined waitlist items */
  rowAfterDivider = row;
  for (var key in declinedItems) {
    removedItemsSheet.getSheets()[0].getRange(row, 1).setValue(key);
    removedItemsSheet.getSheets()[0].getRange(row, 1).setWrap(true);
    removedItemsSheet.getSheets()[0].getRange(row, 3).setValue(declinedItems[key].length);
    removedItemsSheet.getSheets()[0].getRange(row, 1,1,5).setBorder(true, true, true, true, false, true, null, SpreadsheetApp.BorderStyle.SOLID);
    row++;
  }
  /* Check to see if there are values to sort */
  if(Object.keys(declinedItems).length > 0)
    removedItemsSheet.getSheets()[0].getRange(rowAfterDivider,1,(row-rowAfterDivider),5).sort([{column: 3, ascending: false}, {column: 1, ascending: true}]);
}
