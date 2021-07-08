// Written by Andrew Stillman
// Published under GNU General Public License, version 3 (GPL-3.0)
// See restrictions at http://www.opensource.org/licenses/gpl-3.0.html
// Updated and Tweaked by Michael Back

var headers = ['Org Unit Path', 'Annotated Location', 'Annotated Asset ID', 'Serial Number', 'Notes', 'Annotated User', 'Recent Users', 'Status', 'OS Version', 'Last Sync', 'Mac Address', 'Ethernet Mac Address', 'etag', 'Platform Version', 'Device ID', 'Last Enrollment', 'Active Time', 'Model	Firmware Version', 'Boot Mode', 'Support End Date'];

// Not used, furture plan to only show colmns that match these headers
// var visibleHeaders = ['Org Unit Path', 'Annotated Location', 'Annotated Asset ID', 'Serial Number', 'Notes', 'Recent Users', 'Status', 'OS Version', 'Last Sync'];

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Custom Menu')
    .addItem('First Run', 'menuItem1')
    .addSeparator()
    .addSubMenu(ui.createMenu('Sub-menu')
      .addItem('Format Headers', 'menuItem2')
      .addItem('Get Devices', 'menuItem3')
      .addItem('--Blank--', 'menuItem4')
      .addItem('Get 100 Devices', 'menuItem5'))
    .addToUi();
}

//https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet
// The onOpen function is executed automatically every time a Spreadsheet is loaded
// function onOpen() {
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var menuEntries = [];
//   // When the user clicks on "addMenuExample" then "Menu Entry 1", the function function1 is
//   // executed.
//   menuEntries.push({name: "Menu Entry 1", functionName: "function1"});
//   menuEntries.push(null); // line separator
//   menuEntries.push({name: "Menu Entry 2", functionName: "function2"});

//   ss.addMenu("addMenuExample", menuEntries);
// }

function menuItem1() {
  firstRun();
  createSheets();
  // clearSheet('Device Info');
  // clearSheet('Compare');
  setHeader('Device Info');
  setHeader('Compare');
  sanatizeMacInput('Device Info');
  // sanatizeMacInput('Compare');
  hideCompare();
}

function menuItem2() {
  // SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  //   .alert('You clicked the second menu item!');
  // getWidths('Device Info');
  setHeader('Device Info');
  setDetails('Device Info');
  filterSheet('Device Info');
  setHeader('Compare');
  setDetails('Compare');
  filterSheet('Compare');
}

function menuItem3() {
  // SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  //   .alert('You clicked the second menu item!');
  clearSheet('Device Info');
  setHeader('Device Info');
  listChomeDevices();
  setWrap('Device Info');
  setHeader('Device Info');
  filterSheet('Device Info');

  //Now Copy the info to comapre later  
  showCompare();
  clearSheet('Compare');
  copyToCompare();
  setWrap('Compare');
  setHeader('Compare');
  filterSheet('Compare');
  hideCompare();
  Browser.msgBox("Finished getting devices");
}

function menuItem4() {
  // SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  //   .alert('You clicked the second menu item!');
  updateDevices();
}

function menuItem5() {
  // SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  //   .alert('You clicked the second menu item!');
  //------------
  //Testing area
  //------------

  clearSheet('Device Info');
  setHeader('Device Info');
  firstListChomeDevices();
  setWrap('Device Info');
  setHeader('Device Info');
  filterSheet('Device Info');

  //Now Copy the info to comapre later  
  showCompare();
  clearSheet('Compare');
  copyToCompare();
  setWrap('Compare');
  setHeader('Compare');
  filterSheet('Compare');
  hideCompare();
  Browser.msgBox("Finished getting devices");
}

function firstRun() {
  Browser.msgBox("Do not rename the sheets. The script uses the sheets names. \\n If they are changes the script will not work.")
}

function onEdit(e) {
}

function sanatizeMacInput(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // var sheet = ss.getSheets()[0];
  var sheet = ss.getSheetByName(sheetName);
  //Ser MAC to regular text
  sheet.getRange(1, letterToColumn('K'), sheet.getLastRow()).setNumberFormat("@");
}

function getWidths(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // var sheet = ss.getSheets()[0];
  var sheet = ss.getSheetByName(sheetName);
  Browser.msgBox(sheet.getColumnWidth(letterToColumn('A')));
  Browser.msgBox(sheet.getColumnWidth(letterToColumn('B')));
  Browser.msgBox(sheet.getColumnWidth(letterToColumn('C')));
  Browser.msgBox(sheet.getColumnWidth(letterToColumn('D')));
}

function createSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetsCount = ss.getNumSheets();
  var sheets = ss.getSheets();
  try {
    ss.insertSheet('Device Info', 0);
  } catch (e) {
    Logger.log("Sheet already exist.");
    Logger.log(e);
  }
  for (var i = 0; i < sheetsCount; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();
    Logger.log(sheetName);
    if (sheetName != "Device Info") {
      Logger.log("DELETE!" + sheet);
      // ss.deleteSheet(sheet);
    } else {
      Logger.log("No sheets to delete");
    }
  }
  try {
    ss.insertSheet('Compare', 1);
  } catch (e) {
    Logger.log("Sheet already exist.");
    Logger.log(e);
  }
  ss.getSheetByName('Device Info').activate();
}

function copyToCompare() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Compare');
  sheet.showSheet();
  var copyFromSheet = ss.getSheetByName('Device Info');
  //remove all formatting to keep the sheets the same
  try {
    sheet.clearFormats();
    sheet.getFilter().remove();
    copyFromSheet.clearFormats();
    copyFromSheet.getFilter().remove();
  } catch (e) {

  }
  var rangeToCopy = copyFromSheet.getRange(1, 1, copyFromSheet.getMaxRows(), copyFromSheet.getMaxColumns());
  if (sheet == null) {
    Logger.log("Compare Sheet Missing. Adding Sheet Now");
    try {
      ss.insertSheet('Compare', 1);
      ss.getSheetByName('Compare').hideSheet();
    } catch (e) {
      Logger.log("Sheet already exist.");
      Logger.log(e);
    }
  }
  rangeToCopy.copyTo(sheet.getRange(1, 1));
  setHeader('Compare');
  setHeader('Device Info');
  SpreadsheetApp.flush();
}

function showCompare() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Compare');
  try {
    ss.insertSheet('Compare', 1);
    sheet = ss.getSheetByName('Compare');
    sheet.showSheet();
  } catch (e) { }
  sheet.activate();
}
function hideCompare() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Compare');
  sheet.hideSheet();
}

function clearSheet(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // var sheet = ss.getSheets()[0];
  var sheet = ss.getSheetByName(sheetName);
  var maxRow = sheet.getMaxRows();
  var maxColumn = sheet.getMaxColumns();
  try {
    // Clears all content and Formatting
    sheet.clearContents();
    sheet.clearFormats();
    sheet.getFilter().remove();
  } catch (e) { }
  // Make sure it has at least 100 rows
  if (maxRow < 0) {
    sheet.insertRows(maxRow, 2 - maxRow);
  } else if (maxRow == 2) {
    // Do nothing
  } else {
    sheet.deleteRows(2, maxRow - 2);
  }
  // Makes sure it has all headers and one free space
  if (maxColumn < letterToColumn('U')) {
    sheet.insertColumns(maxColumn, letterToColumn('U') - maxColumn);
  } else if (maxColumn == letterToColumn('U')) {
    // Do nothing
  } else {
    sheet.deleteColumns(letterToColumn('U'), maxColumn - letterToColumn('U'));
  }
  SpreadsheetApp.flush();
}

function setHeader(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // var sheet = ss.getSheets()[0];
  var sheet = ss.getSheetByName(sheetName);
  sheet.clearFormats();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, letterToColumn('F')).setFontWeight('bold').setHorizontalAlignment("center").setBackground('#74b9ff');
  sheet.getRange(1, letterToColumn('F') + 1, 1, headers.length - letterToColumn('D')).setFontWeight('bold').setHorizontalAlignment("center").setBackground('grey');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(letterToColumn('F'));

  sheet.setColumnWidths(1, headers.length, 100);
  sheet.setColumnWidths(1, 1, 360);
  sheet.setColumnWidth(letterToColumn('B'), 145);
  sheet.setColumnWidth(letterToColumn('C'), 130);
  sheet.setColumnWidth(letterToColumn('D'), 110);
  sheet.setColumnWidth(letterToColumn('T'), 120);

  // Hides unneed columns
  //Want to show the first 9 and hide the rest
  sheet.hideColumns(9, headers.length - 9);

  Logger.log("Set Headers for sheet: " + sheetName);

  SpreadsheetApp.flush();
}

function setDetails(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  sheet.setFrozenColumns(letterToColumn('F'));
  SpreadsheetApp.flush();
}

function setWrap(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // var sheet = ss.getSheets()[0];
  var sheet = ss.getSheetByName(sheetName);
  var maxRow = sheet.getMaxRows();
  var maxColumn = sheet.getMaxColumns();
  sheet.getRange(1, 1, maxRow, maxColumn).setWrap(false);
  Logger.log("Wrap Set to false for sheet: " + sheetName);
  SpreadsheetApp.flush();
}

function filterSheet(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // var sheet = ss.getSheets()[0];
  var sheet = ss.getSheetByName(sheetName);
  var maxRow = sheet.getMaxRows();
  // Updates any current filters
  try {
    //Remove Filters
    sheet.getFilter().remove();
  } catch (e) {
    Logger.log("No Filters to remove.");
  }
  //Hard coded the cell for creating the filter
  sheet.getRange('A1:A' + maxRow).createFilter();
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

function updateDevices() {
  var ok = Browser.msgBox('Are you sure?  This will update the Organizational Unit, Annotated User, Annotated Location, and Notes for all devices listed in the sheet', Browser.Buttons.OK_CANCEL);
  Browser.msgBox("Please wait untill another box pops up before changing anyhting or closing the tab.")
  if (ok == "ok") {
    try {
      //https://developers.google.com/apps-script/articles/mail_merge
      //https://stackoverflow.com/questions/45987095/apps-script-getrowsdata-function-deprecated
      var updatedCount = 0;
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName('Device Info');
      var compareTo = ss.getSheetByName('Compare');
      sanatizeMacInput('Device Info');
      sanatizeMacInput('Compare');
      compareTo.showSheet();
      var updateFailed = false;
      if (sheet.getLastRow() > 1) {
        var data = getRowsData(sheet);
        var compareData = getRowsData(compareTo);
        for (var i = 0; i < data.length; i++) {
          if (data[i].status === "ACTIVE") {
            if (data[i].orgUnitPath != compareData[i].orgUnitPath ||
              data[i].annotatedLocation != compareData[i].annotatedLocation ||
              data[i].notes != compareData[i].notes ||
              data[i].annotatedUser != compareData[i].annotatedUser ||
              data[i].annotatedAssetId != compareData[i].annotatedAssetId
            ) {
              Logger.log("Loop Number : " + i + "\n" +
                "Row Number : " + (i + 2) + "\n" +
                data[i].orgUnitPath + ':' + compareData[i].orgUnitPath + "\n" +
                data[i].annotatedLocation + ':' + compareData[i].annotatedLocation + "\n" +
                data[i].notes + ':' + compareData[i].notes + "\n" +
                data[i].annotatedUser + ':' + compareData[i].annotatedUser + "\n" +
                data[i].annotatedAssetId + ':' + compareData[i].annotatedAssetId);
              //Must set Recent Users to null beucase it is expecting an object, not a single user like on the sheet
              data[i].recentUsers = null;
              try {
                AdminDirectory.Chromeosdevices.update(data[i], 'my_customer', data[i].deviceId);
                Logger.log("At: " + i + data[i], 'my_customer', data[i].deviceId);
                updatedCount++;
              } catch (e) {
                Logger.log("AdminDirectory error at: " + i + "\nError Msg: " + e);
                updateFailed = true;
                continue;
              }
            }
          }
        }
      }
      if (updateFailed) {
        Browser.msgBox("AdminDirectory update error. \\nCheck Logs.");
      } else {
        Browser.msgBox(updatedCount + " Chrome devices were updated in the inventory...");
        //After Testing Remove the comments!!!
        if (updatedCount > 0) {
          copyToCompare();
          setHeader('Device Info');
          setHeader('Compare');
          compareTo.hideSheet();
        }
      }
    } catch (err) {
      Browser.msgBox(err.message);
    }
  }
}

function getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(range.getValues(), normalizeHeaders(headers));
}
