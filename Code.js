// Written by Andrew Stillman
// Published under GNU General Public License, version 3 (GPL-3.0)
// See restrictions at http://www.opensource.org/licenses/gpl-3.0.html
// Updated and Tweaked by Michael Back

var headers = ['Org Unit Path', 'Annotated Location', 'Annotated Asset ID', 'Serial Number', 'Notes', 'Recent Users', 'Status', 'OS Version', 'Last Sync', 'Annotated User', 'Mac Address', 'Ethernet Mac Address', 'etag', 'Platform Version', 'Device ID', 'Last Enrollment', 'Active Time', 'Model	Firmware Version', 'Boot Mode', 'Support End Date'];

// Not used, furture plan to only show colmns that match these headers
var visibleHeaders = ['Org Unit Path', 'Annotated Location', 'Annotated Asset ID', 'Serial Number', 'Notes', 'Recent Users', 'Status', 'OS Version', 'Last Sync'];

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Custom Menu')
    .addItem('Clear Sheet & Rows', 'menuItem1')
    .addSeparator()
    .addSubMenu(ui.createMenu('Sub-menu')
      .addItem('Format Headers', 'menuItem2')
      .addItem('Get Devices', 'menuItem3')
      .addItem('--Blank--', 'menuItem4')
      .addItem('Get 100 Devices', 'menuItem5'))
    .addToUi();
}

function menuItem1() {
  firstRun();
}

function menuItem2() {
  // SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  //   .alert('You clicked the second menu item!');
  // getWidths();
  filterSheet();
  setHeader();
  setDetails();
}

function menuItem3() {
  // SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  //   .alert('You clicked the second menu item!');
  clearSheet();
  setHeader();
  listChomeDevices();
  setWrap();
  setHeader();
  filterSheet();
}

function menuItem4() {
  // SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  //   .alert('You clicked the second menu item!');
  updateDevices();
}

function menuItem5() {
  // SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  //   .alert('You clicked the second menu item!');
  clearSheet();
  setHeader();
  firstListChomeDevices();
  setWrap();
  setHeader();
  filterSheet();
}

function firstRun() {
}

function onEdit(e) {
}

function getWidths() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  Browser.msgBox(sheet.getColumnWidth(letterToColumn('A')));
  Browser.msgBox(sheet.getColumnWidth(letterToColumn('B')));
  Browser.msgBox(sheet.getColumnWidth(letterToColumn('C')));
  Browser.msgBox(sheet.getColumnWidth(letterToColumn('D')));
}

function clearSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var maxRow = sheet.getMaxRows();
  var maxColumn = sheet.getMaxColumns();
  try {
    // Clears all content and Formatting
    sheet.clearContents();
    sheet.clearFormats();
    ss.getActiveSheet().getFilter().remove();
  } catch (e) {}
  // Make sure it has at least 100 rows
  if (maxRow < 0) {
    sheet.insertRows(maxRow, 2 - maxRow);
  } else if (maxRow == 2) {
    // Do nothing
  } else {
    sheet.deleteRows(2, maxRow - 2);
  }
  // Makes sure it has all headers and one free space
  if (maxColumn < letterToColumn('U')) {sheet.insertColumns(maxColumn, letterToColumn('U') - maxColumn);
  } else if (maxColumn == letterToColumn('U')) {
    // Do nothing
  } else {sheet.deleteColumns(letterToColumn('U'), maxColumn - letterToColumn('U'));
  }
  sheet.setColumnWidths(1, headers.length, 100);
  sheet.setColumnWidths(1, 1, 360);
  sheet.setColumnWidth(letterToColumn('B'), 145);
  sheet.setColumnWidth(letterToColumn('C'), 130);
  sheet.setColumnWidth(letterToColumn('D'), 110);

  // Hides unneed columns
  //Want to show the first 9 and hide the rest
  sheet.hideColumns(9, headers.length - 9);
  SpreadsheetApp.flush();
}

function setHeader() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  sheet.clearFormats();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, letterToColumn('D')).setFontWeight('bold').setHorizontalAlignment("center").setBackground('#74b9ff');
  sheet.getRange(1, letterToColumn('D') + 1, 1, headers.length - letterToColumn('D')).setFontWeight('bold').setHorizontalAlignment("center").setBackground('grey');
  sheet.setFrozenRows(1);
  SpreadsheetApp.flush();
}

function setDetails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  sheet.setFrozenColumns(letterToColumn('D'));
  SpreadsheetApp.flush();
}

function setWrap() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var maxRow = sheet.getMaxRows();
  var maxColumn = sheet.getMaxColumns();
  sheet.getRange(1, 1, maxRow, maxColumn).setWrap(false);
  SpreadsheetApp.flush();
}

function filterSheet() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var maxRow = sheet.getMaxRows();
  // Updates any current filters
  try {
    //Remove Filters
    ss.getActiveSheet().getFilter().remove();
    //Hard coded the cell for creating the filter
    spreadsheet.getRange('A1:A' + maxRow).createFilter().sort(letterToColumn('A'), true);
  } catch (e) { }
}


function listChomeDevices() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    var allDevices = [];
    SpreadsheetApp.flush();
    // This grabs the first 100 devices and then will allow 
    // the while loop to go throught the rest becuase of response.nextPageToken
    var response = AdminDirectory.Chromeosdevices.list('my_customer', { maxResults: 100, projection: "FULL" });
    allDevices = allDevices.concat(response.chromeosdevices);
    // Browser.msgBox(Object.entries(allDevices));
    // This grabs all the devices (as long as it has a nextPageToken)
    while (response.nextPageToken) {
      response = AdminDirectory.Chromeosdevices.list('my_customer', { maxResults: 100, projection: "FULL", pageToken: response.nextPageToken });
      allDevices = allDevices.concat(response.chromeosdevices);
    }
    //https://stackoverflow.com/questions/1078118/how-do-i-iterate-over-a-json-structure
    // https://www.freecodecamp.org/news/javascript-foreach-how-to-loop-through-an-array-in-js/
    // https://zetcode.com/javascript/jsonforeach/
    // This just fills in all the data from allDevices 
    // The the flush is like telling the sheet to refresh to show changes
    Logger.log(allDevices);
    setRowsData(sheet, allDevices);
    SpreadsheetApp.flush();
  } catch (err) {
    Browser.msgBox("Error: " + err.message);
  }
}

function firstListChomeDevices() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    var allDevices = [];
    SpreadsheetApp.flush();
    var response = AdminDirectory.Chromeosdevices.list('my_customer', { maxResults: 100, projection: "FULL" });
    allDevices = allDevices.concat(response.chromeosdevices);
    Logger.log(allDevices);
    setRowsData(sheet, allDevices);
    SpreadsheetApp.flush();
  } catch (err) {
    Browser.msgBox("Error: " + err.message);
  }
}

function updateDevices(){
  
}
