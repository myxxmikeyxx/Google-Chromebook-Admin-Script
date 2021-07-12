// Written by Andrew Stillman
// Published under GNU General Public License, version 3 (GPL-3.0)
// See restrictions at http://www.opensource.org/licenses/gpl-3.0.html
// Updated and Tweaked by Michael Back
//
// https://developers.google.com/admin-sdk/directory/reference/rest/v1/chromeosdevices#ChromeOsDevice
// Maybe add
//https://stackoverflow.com/questions/58064351/is-there-a-way-to-run-two-functions-at-the-same-timesimultaneously-asynchrounou

var headers = ['Org Unit Path', 'Annotated Location', 'Annotated Asset ID', 'Serial Number', 'Notes', 'Annotated User', 'Recent Users', 'Status', 'OS Version', 'Last Sync', 'Mac Address', 'Ethernet Mac Address', 'etag', 'Platform Version', 'Device ID', 'Last Enrollment', 'Active Time', 'Model	Firmware Version', 'Boot Mode', 'Support End Date'];

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Chrome Device Management')
    .addItem('First Run', 'menuItem1')
    .addSeparator()
    .addSubMenu(ui.createMenu('Get Devices')
      .addItem('Get Devices', 'menuItem3'))
    .addSubMenu(ui.createMenu('Update Devices')
      .addItem('Update Device Info', 'menuItem4'))
    .addSubMenu(ui.createMenu('Restore Backup')
      .addItem('Restore Backup', 'menuItem9'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Testing Zone')
      .addSeparator()
      .addItem('Do Not Click Anything', 'showcompare')
      .addItem('For Testing Only', 'hideSheet("Compare")')
      .addSeparator()
      .addItem('Format Headers', 'menuItem2')
      .addItem('Get 100 Devices', 'menuItem5')
      .addItem('Data Validation', 'menuItem6')
      .addItem('Filter Testing', 'menuItem7')
      .addItem('Remove All Sheets', 'menuItem8'))
    .addToUi();
}

// function onOpen() {
//   var ui = SpreadsheetApp.getUi();
//   // Or DocumentApp or FormApp.
//   ui.createMenu('Custom Menu')
//     .addItem('First Run', 'menuItem1')
//     .addSeparator()
//     .addSubMenu(ui.createMenu('Manage Devices')
//       .addItem('Get Devices', 'menuItem3')
//       .addItem('Update Device Info', 'menuItem4'))
//       .addSeparator()
//     .addSubMenu(ui.createMenu('--Testing Zone--')
//       .addSeparator()
//       .addItem('Do Not Click Anything', '')
//       .addItem('For Testing Only', '')
//       .addSeparator()
//       .addItem('Format Headers', 'menuItem2')
//       .addItem('Get 100 Devices', 'menuItem5')
//       .addItem('Data Validation', 'menuItem6')
//       .addItem('Filter Testing', 'menuItem7')
//       .addItem('Remove All Sheets', 'menuItem8'))
//     .addToUi();
// }

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
  // filterTesting();
  firstRun();
  createSheets();
  var ok = Browser.msgBox('Do you want to clear the sheets? If not click anything other than OK. \\n\\n This will not clear Useful Formulas content.', Browser.Buttons.OK_CANCEL);
  if (ok == "ok") {
    clearSheet('Device Info');
    clearSheet('Compare');
  }
  setHeader('Device Info');
  setHeader('Compare');
  filterSheet('Device Info');
  filterSheet('Compare');
  dataVal('Device Info');
  dataVal('Compare');
  hideSheet('Compare');
}

function menuItem2() {
  // SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  //   .alert('You clicked the second menu item!');
  // getWidths('Device Info');
  setHeader('Device Info');
  setDetails('Device Info');
  filterSheet('Device Info');
  dataVal('Device Info');
  setHeader('Compare');
  setDetails('Compare');
  filterSheet('Compare');
  dataVal('Compare');
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
  dataVal('Device Info');

  //Now Copy the info to compare sheet  
  showCompare();
  clearSheet('Compare');
  copyToSheet('Device Info', 'Compare');
  setWrap('Compare');
  setHeader('Compare');
  filterSheet('Compare');
  dataVal('Device Info');
  hideSheet('Compare');
  Browser.msgBox("Finished getting devices");
}

function menuItem4() {
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
  dataVal('Device Info');

  //Now Copy the info to comapre later  
  showCompare();
  clearSheet('Compare');
  copyToSheet('Device Info', 'Compare');
  setWrap('Compare');
  setHeader('Compare');
  filterSheet('Compare');
  dataVal('Compare');
  hideSheet('Compare');
  Browser.msgBox("Finished getting devices");
}

function menuItem6() {
  dataVal('Device Info');
}
function menuItem7() {
  filterTesting();
}

function menuItem8() {
  clearAllSheets();
}

function menuItem9() {
  restoreDevices();
}

function firstRun() {
  Browser.msgBox("User must have access to google admin and ability to manage chrome devices." +
    "\\nDo not rename the sheets. The script uses the sheets names. \\n If they are changes the script will not work.");
  Browser.msgBox("This script should only show 'ACTIVE' Devices.");
  Browser.msgBox("Get Devices to update the list of devices before making any changes. It should only change devices that the information if different on the Compare sheet," +
    "\\nMeaning if people are in admin chainging items it should not change that information unless you are changing it on the sheet as well, then where is the most recent save/push will be kept.")
  // Browser.msgBox("Lastly if a box is blank it is marked as undefined. \\nThis means that it will not change the information in google admin. \\n!!!YOU MUST PUT A SPACE IN THE SPOT TO MAKE IT BLANK IN ADMIN!!!");
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
    Logger.log("Device Info sheet already exist.");
    Logger.log(e);
  }
  for (var i = 0; i < sheetsCount; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();
    Logger.log(sheetName);
    if (sheetName != "Device Info") {
      if (sheetName != "For Work") {
        if (sheetName != "Backup") {
          if (sheetName != "Useful Formulas") {
            Logger.log("DELETE! " + sheet);
            ss.deleteSheet(sheet);
          }
        }
      }
    } else {
      Logger.log("No sheets to delete.");
    }
  }
  try {
    ss.insertSheet('Compare', 1);
    hideSheet('Compare');
  } catch (e) {
    hideSheet('Compare');
    Logger.log("Compare sheet already exist.");
    Logger.log(e);
  }
  try {
    ss.insertSheet('For Work', 2);
  } catch (e) {
    Logger.log("For Work sheet already exist.");
    Logger.log(e);
  }
  try {
    ss.insertSheet('Backup', 3);
    hideSheet('Backup');
  } catch (e) {
    hideSheet('Backup');
    Logger.log("Backup sheet already exist.");
    Logger.log(e);
  }
  try {
    ss.insertSheet('Useful Formulas', 4);
    ss.getSheetByName('Useful Formulas');
    ss.getRange('A1').setValue("\'=IF(ISNA(VLOOKUP(D39,'For Work'!A:K,6, false)),\"\", VLOOKUP(D39,'For Work'!A:K,6, false))");
    ss.getRange('A2').setValue("\'=IF(ISNA(VLOOKUP(D3,'For Work'!A:K,2, false)),\"\", VLOOKUP(D3,'For Work'!A:K,2, false))");
    //make an array of formuals that are useful and do a for loop to add them to the sheet
  } catch (e) {
    Logger.log("Useful Formulas sheet already exist.");
    Logger.log(e);
  }
  ss.getSheetByName('Device Info').activate();
}

function clearAllSheets() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetsCount = ss.getNumSheets();
  var sheets = ss.getSheets();
  try {
    ss.insertSheet('Sheet 1', 0);
  } catch (e) {
    Logger.log("Sheet 1 already exist.");
    Logger.log(e);
  }
  for (var i = 0; i < sheetsCount; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();
    Logger.log(sheetName);
    if (sheetName != "Sheet 1") {
      Logger.log("DELETE! " + sheet);
      ss.deleteSheet(sheet);
    } else {
      Logger.log("No sheets to delete.");
    }
  }
}

function copyToSheet(sheetName, copyTo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(copyTo);
  sheet.showSheet();
  var copyFromSheet = ss.getSheetByName(sheetName);
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
function hideSheet(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
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
  SpreadsheetApp.flush();
}

function dataVal(sheetName) {
  //https://stackoverflow.com/questions/59216381/google-script-retrieving-default-values-for-filter
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  // Set the data validation for cell A2 to require a value from A2:A (lastrow).
  var cell = sheet.getRange('A2:A' + sheet.getLastRow());
  var range = sheet.getRange('A2:A' + sheet.getLastRow());
  var rule = SpreadsheetApp.newDataValidation().requireValueInRange(range).build();
  cell.clearDataValidations();
  cell.setDataValidation(rule);
  // https://developers.google.com/apps-script/reference/spreadsheet/data-validation
  // https://developers.google.com/apps-script/reference/spreadsheet/data-validation-builder

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
  Browser.msgBox("After closing this, please wait until another box pops up after this one, \n before changing anything or closing the tab.")
  if (ok == "ok") {
    try {
      //https://developers.google.com/apps-script/articles/mail_merge
      //https://stackoverflow.com/questions/45987095/apps-script-getrowsdata-function-deprecated
      var updatedCount = 0;
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName('Device Info');
      var compareTo = ss.getSheetByName('Compare');
      var filter = sheet.getFilter();
      try {
        ss.getSheetByName('Device Info').getRange('A1').clearDataValidations();
        ss.getSheetByName('Compare').getRange('A1').clearDataValidations();
        ss.getSheetByName('Device Info').getFilter().remove();
        ss.getSheetByName('Compare').getFilter().remove();
      } catch (e) {
        Logger.log("Filter already removed.");
      }
      try {
        sanatizeMacInput('Device Info');
        sanatizeMacInput('Compare');
        sanatizeMacInput('Backup');
      } catch (e) {
        Logger.log("Backup Sheet is empty.");
      }
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
              // Logger.log("Loop Number : " + i + "\n" +
              //   "Row Number : " + (i + 2) + "\n" +
              //   data[i].orgUnitPath + ':' + compareData[i].orgUnitPath + "\n" +
              //   data[i].annotatedLocation + ':' + compareData[i].annotatedLocation + "\n" +
              //   data[i].notes + ':' + compareData[i].notes + "\n" +
              //   data[i].annotatedUser + ':' + compareData[i].annotatedUser + "\n" +
              //   data[i].annotatedAssetId + ':' + compareData[i].annotatedAssetId);
              //Sets Recent Users to null becuase it will cause problems if it's not an object
              data[i].recentUsers = null;
              try {
                //https://developers.google.com/admin-sdk/directory/reference/rest/v1/chromeosdevices#ChromeOsDevice
                if (data[i].annotatedAssetId == null) {
                  data[i].annotatedAssetId = '';
                }
                if (data[i].annotatedLocation == null) {
                  data[i].annotatedLocation = '';
                }
                if (data[i].annotatedUser == null) {
                  data[i].annotatedUser = '';
                }
                if (data[i].notes == null) {
                  data[i].notes = '';
                }
                AdminDirectory.Chromeosdevices.update(data[i], 'my_customer', data[i].deviceId);
                //Logger.log("At: " + i + data[i], 'my_customer', data[i].deviceId);
                updatedCount++;
              } catch (e) {
                Logger.log("AdminDirectory error at row: " + (i + 2) + "\nError Msg: " + e);
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
        if (updatedCount = 1) {
          Browser.msgBox(updatedCount + " Chrome device was updated in the inventory...");
          Logger.log(updatedCount + " Chrome device was updated in the inventory...");
        } else {
          Browser.msgBox(updatedCount + " Chrome devices were updated in the inventory...");
          Logger.log(updatedCount + " Chrome devices were updated in the inventory...");
        }
        if (updatedCount >= 0) {
          //Makes a Backup
          copyToSheet('Compare', 'Backup');
          //Updates compare so next update it saves time and tries to update only changed device info
          copyToSheet('Device Info', 'Compare');
          setHeader('Device Info');
          setHeader('Compare');
          //Applys back the filtered view the user 
          sheet.getRange('A1:A' + sheet.getLastRow()).createFilter().setColumnFilterCriteria(1, filter);
          filterSheet('Compare');
          dataVal('Device Info');
          dataVal('Compare');
          hideSheet('Compare');
          hideSheet('Backup');
        }
      }
    } catch (err) {
      Browser.msgBox(err.message);
    }
  }
}

function filterTesting() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Device Info');
  var filter = sheet.getFilter().getColumnFilterCriteria(1);
  sheet.getFilter().remove();
  sheet.getRange('A1:A' + sheet.getLastRow()).createFilter().setColumnFilterCriteria(1, filter);
}

function getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(range.getValues(), normalizeHeaders(headers));
}

function restoreDevices() {
  var ok = Browser.msgBox('Are you sure you want to restore from backup?  This will update the Organizational Unit, Annotated User, Annotated Location, and Notes for all devices back to before the last push.', Browser.Buttons.OK_CANCEL);
  Browser.msgBox("After closing this, please wait until another box pops up after this one, \n before changing anything or closing the tab.")
  if (isSheetEmpty('Backup') != "") {
    if (ok == "ok") {
      try {
        var updatedCount = 0;
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName('Backup');
        try {
          ss.getSheetByName('Backup').getRange('A1').clearDataValidations();
          ss.getSheetByName('Backup').getFilter().remove();
        } catch (e) {
          Logger.log("Filter already removed on Backup Sheet.");
        }
        try {
          sanatizeMacInput('Backup');
        } catch (e) {
        }
        var updateFailed = false;
        if (sheet.getLastRow() > 1) {
          var data = getRowsData(sheet);
          for (var i = 0; i < data.length; i++) {
            if (data[i].status === "ACTIVE") {
              //Sets Recent Users to null because it will cause problems if it's not an object
              data[i].recentUsers = null;
              try {
                if (data[i].annotatedAssetId == null) {
                  data[i].annotatedAssetId = '';
                }
                if (data[i].annotatedLocation == null) {
                  data[i].annotatedLocation = '';
                }
                if (data[i].annotatedUser == null) {
                  data[i].annotatedUser = '';
                }
                if (data[i].notes == null) {
                  data[i].notes = '';
                }
                AdminDirectory.Chromeosdevices.update(data[i], 'my_customer', data[i].deviceId);
                updatedCount++;
              } catch (e) {
                Logger.log("AdminDirectory error at row: " + (i + 2) + "\nError Msg: " + e);
                updateFailed = true;
                continue;
              }
            }
          }
        }
        if (updateFailed) {
          Browser.msgBox("AdminDirectory update error. \\nCheck Logs.");
        } else {
          if (updatedCount = 1) {
            Browser.msgBox(updatedCount + " Chrome device was updated in the inventory...");
            Logger.log(updatedCount + " Chrome device was updated in the inventory...");
          } else {
            Browser.msgBox(updatedCount + " Chrome devices were updated in the inventory...");
            Logger.log(updatedCount + " Chrome devices were updated in the inventory...");
          }
          if (updatedCount >= 0) {
            hideSheet('Backup');
          }
        }
      } catch (err) {
        Browser.msgBox(err.message);
      }
    }
  } else {
    Logger.log("Backup Sheet is empty.");
    Browser.msgBox("Backup Sheet is empty.");
  }
}

function isSheetEmpty(sheet) {
  return sheet.getDataRange().getValues().join("") === "";
}
