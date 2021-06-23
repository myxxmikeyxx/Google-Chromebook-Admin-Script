// Written by Andrew Stillman for New Visions for Public Schools
// Published under GNU General Public License, version 3 (GPL-3.0)
// See restrictions at http://www.opensource.org/licenses/gpl-3.0.html


var headers = ['etag', 'Org Unit Path', 'Serial Number', 'Platform Version', 'Device Id', 'Status', 'Last Enrollment Time',
  'Firmware Version', 'Last Sync', 'OS Version', 'Boot Mode', 'Annotated Location', 'Notes', 'Annotated User', 'Mac Address', 'ANYTHING PUT HERE WILL DO NOTHING'];
var specialauthId = '0B7-FEGXAo-DGVklteGtFT2trOFU';
var image1Id = '0B7-FEGXAo-DGZ2txTWFnZ05SVVU';
var image2Id = '0B7-FEGXAo-DGeTUydWNHTjVDUVE';
var image3Id = '0B7-FEGXAo-DGRFNxX054S3p4QUE';
var imageBase = 'https://drive.google.com/uc?export=download&id=';



function myFunction() {

}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Custom Menu')
    .addItem('Clear Sheet & Rows', 'menuItem1')
    .addSeparator()
    .addSubMenu(ui.createMenu('Sub-menu')
      .addItem('Format Headers', 'menuItem2')
      .addItem('Get Devices', 'menuItem3')
      .addItem('Get Devices Testing', 'menuItem4')
      .addItem('Test Filter', 'menuItem5'))
    .addToUi();
}

function menuItem1() {
  // SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  //   .alert('You clicked the first menu item!');
  clearSheet();
}

function menuItem2() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
    .alert('You clicked the second menu item!');
  setHeader();
}

function menuItem3() {
  // SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  //   .alert('You clicked the second menu item!');
  listChomeDevices();
}

function menuItem4() {
  // SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  //   .alert('You clicked the second menu item!');
  testListChomeDevices();
}

function menuItem5() {
  // SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  //   .alert('You clicked the second menu item!');
  filterSheet();
}

function onEdit(e) {
  // var sheetToWatch= 'Devices Checked out & Returns',
  // columnToWatch = 2, columnToStamp = 1;
  //     if (e.range.columnStart !== columnToWatch || e.source.getActiveSheet()
  //         .getName() !== sheetToWatch || !e.value) return;
  //     e.source.getActiveSheet()
  //         .getRange(e.range.rowStart, columnToStamp)
  //         .setValue(new Date());
  // columnToWatch = 2, columnToStamp = 12;
  //     if (e.range.columnStart !== columnToWatch || e.source.getActiveSheet()
  //         .getName() !== sheetToWatch || !e.value) return;
  //     e.source.getActiveSheet()
  //         .getRange(e.range.rowStart, columnToStamp)
  //     .setFormula("=B3+4");
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
  } catch (e) {
  }

  // Make sure it has at least 100 rows
  if (maxRow < 100) {
    sheet.insertRows(maxRow, 100 - maxRow);
  } else if (maxRow == 100) {
    // Do nothing
  } else {
    sheet.deleteRows(100, maxRow - 100);
  }
  // Makes sure it has all headers and one free space
  if (maxColumn < headers.length) {
    sheet.insertColumns(maxColumn, headers.length - maxColumn);
  } else if (maxColumn == headers.length) {
    // Do nothing
  } else {
    sheet.deleteColumns(headers.length, maxColumn - headers.length);
  }


  sheet.setColumnWidths(1, headers.length, 200);

  // Hides unneed columns
  sheet.hideColumns(letterToColumn('A'));
  sheet.hideColumns(letterToColumn('D'));
  sheet.hideColumns(letterToColumn('E'));
  sheet.hideColumns(letterToColumn('F'));
  sheet.hideColumns(letterToColumn('G'));
  sheet.hideColumns(letterToColumn('H'));
  sheet.hideColumns(letterToColumn('I'));
  sheet.hideColumns(letterToColumn('J'));
  sheet.hideColumns(letterToColumn('K'));
  // sheet.hideColumns(letterToColumn('L'));
  // sheet.hideColumns(letterToColumn('M'));
  sheet.hideColumns(letterToColumn('N'));
  sheet.hideColumns(letterToColumn('O'));
  SpreadsheetApp.flush();
}

function setHeader() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  sheet.clearFormats();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setHorizontalAlignment("center");
  sheet.setFrozenRows(1);
  SpreadsheetApp.flush();
}


function listChomeDevices() {
  ScriptProperties.setProperty('adminSDKAuthorized', 'true');
  var ss = SpreadsheetApp.getActiveSpreadsheet().getId();
  var sss = SpreadsheetApp.openById(ss).getActiveSheet();
  var optionalArgs = {
    maxResults: 100,
    orderBy: 'SERIAL_NUMBER',
    projection: 'BASIC'
  };
  var response = AdminDirectory.Chromeosdevices.list('my_customer', optionalArgs);
  var devices = response.chromeosdevices;
  if (devices && devices.length > 0) {
    Logger.log('Devices:');
    for (i = 0; i < devices.length; i++) {
      var device = devices[i];
      sss.getRange(i + 2, 1).setValue(device.serialNumber);
      sss.getRange(i + 2, 2).setValue(device.annotatedAssetId);
      sss.getRange(i + 2, 3).setValue(device.model);
    }
  } else {
    Logger.log('No users found.');
  }
}
function testListChomeDevices() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    var allDevices = [];
    var maxRow = sheet.getMaxRows();
    sheet.clearContents();
    if (maxRow > 100) {
      clearSheet();
    }
    SpreadsheetApp.flush();

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.autoResizeColumn(headers.length);
    SpreadsheetApp.flush();

    var response = AdminDirectory.Chromeosdevices.list('my_customer', { maxResults: 100, projection: "FULL" });
    allDevices = allDevices.concat(response.chromeosdevices);
    while (response.nextPageToken) {
      response = AdminDirectory.Chromeosdevices.list('my_customer', { maxResults: 100, projection: "FULL", pageToken: response.nextPageToken });
      allDevices = allDevices.concat(response.chromeosdevices);
    }

    if (allDevices.length > 0) {
      if (allDevices.length > maxRow) {
        sheet.insertRows(maxRow, (allDevices.length - maxRow));
      }
      setRowsData(sheet, allDevices);
    }
    SpreadsheetApp.flush();
  } catch (err) {
    Browser.msgBox(err.message);
  }
}

function filterSheet() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var maxColumn = sheet.getMaxColumns();
  var maxRow = sheet.getMaxRows();

  // Removes any current filters
  try {spreadsheet.getActiveSheet().getFilter().remove();} catch (e) { }

  // Creates a Filter View of all values in the "Org Unit Path" (it gets rid of duplicates for us, very helpful)
  // spreadsheet.getRange('B1:' + columnToLetter(maxColumn) + maxRow).createFilter();
  spreadsheet.getRange('B1:B' + maxRow).createFilter();
}

