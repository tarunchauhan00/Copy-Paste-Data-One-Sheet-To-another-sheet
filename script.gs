function onOpen(e){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Paste To Another Sheet')
  .addItem('Paste by Range', 'askRange')
  .addItem('Copy All the data', 'syncData')
  .addToUi();
}


 function runsies(){
  importRange(
    "1PhB44GoD_zW-gmErZ2zDgXZgDiGuM_d7r04lqN7L4Zg",  //ID of the source SpreadSheet
    "source!A1:G",           //range i.e. how much have to copy
    "1u1QNoHjxX8G-Ux7GEZek_838POdPX3v4EtA4xjSp_Lw",  //id of the Destination SpreadSheet
    "destination!a1"         //startin range of destination sheet
  );
 }

function importRange(sourceID, sourceRange, destinationID, destinationRangeStart){
  const sourceSS = SpreadsheetApp.openById(sourceID);
  const sourceRng = sourceSS.getRange(sourceRange);
  const sourceVals = sourceRng.getValues();

  const destinationSS = SpreadsheetApp.openById(destinationID);
  const destStartRange = destinationSS.getRange(destinationRangeStart);
  const destSheet = destinationSS.getSheetByName(destStartRange.getSheet().getName());

  destSheet.clear();

  const destRange = destSheet.getRange(
    destStartRange.getRow(),
    destStartRange.getColumn(),
    sourceVals.length,
    sourceVals[0].length

  );

  destRange.setValues(sourceVals);
}


function askRange() {
  // Ask the user for input
  const sourceID = Browser.inputBox("Enter the ID of the source spreadsheet:");
  const sourceRange = Browser.inputBox("Enter the source range (e.g., source!A1:G):");
  const destinationID = Browser.inputBox("Enter the ID of the destination spreadsheet:");
  const destinationRangeStart = Browser.inputBox("Enter the destination start range (e.g., destination!A1):");

  // Call the importRange function with user-provided input
  importRange(sourceID, sourceRange, destinationID, destinationRangeStart);
}

function importRange(sourceID, sourceRange, destinationID, destinationRangeStart) {
  const sourceSS = SpreadsheetApp.openById(sourceID);
  const sourceRng = sourceSS.getRange(sourceRange);
  const sourceVals = sourceRng.getValues();

  const destinationSS = SpreadsheetApp.openById(destinationID);
  const destStartRange = destinationSS.getRange(destinationRangeStart);
  const destSheet = destinationSS.getSheetByName(destStartRange.getSheet().getName());

  destSheet.clear();

  const destRange = destSheet.getRange(
    destStartRange.getRow(),
    destStartRange.getColumn(),
    sourceVals.length,
    sourceVals[0].length
  );

  destRange.setValues(sourceVals);
}







function syncData() {
  // Prompt user for Source Spreadsheet ID and Sheet Name
  var sourceSpreadsheetId = Browser.inputBox('Enter Source Spreadsheet ID');
  var sourceSheetName = Browser.inputBox('Enter Source Sheet Name');

  // Open Source Spreadsheet
  var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
  var sourceData = sourceSheet.getDataRange().getValues();

  // Prompt user for Destination Spreadsheet ID and Sheet Name
  var destSpreadsheetId = Browser.inputBox('Enter Destination Spreadsheet ID');
  var destSheetName = Browser.inputBox('Enter Destination Sheet Name');

  // Open Destination Spreadsheet
  var destSpreadsheet = SpreadsheetApp.openById(destSpreadsheetId);
  var destSheet = destSpreadsheet.getSheetByName(destSheetName);

  // Clear existing data in the destination sheet
  destSheet.clear();

  // Write data to destination sheet
  destSheet.getRange(1, 1, sourceData.length, sourceData[0].length).setValues(sourceData);

  Browser.msgBox('Data synced successfully!');
}

