function getSheetName() {
  // Get the last known sheet name from Script Properties
  var lastSheetName = PropertiesService.getScriptProperties().getProperty('backupSheetName');

  // Ask for the name of the sheet to backup or if you want to use the last known name
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Enter the name of the sheet to backup, or leave it blank to use the last backed up sheet:', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() === ui.Button.CANCEL) {
    // User clicked "Cancel," so abort backup
    ui.alert('Backup aborted.');
    return;
  }

  var sheetNameToBackup = response.getResponseText().trim();
  if (sheetNameToBackup === "") {
    // If the sheet name is blank, use the last known sheet name
    sheetNameToBackup = lastSheetName;
  } else {
    // Save the new sheet name to Script Properties
    PropertiesService.getScriptProperties().setProperty('backupSheetName', sheetNameToBackup);
  }

  // Call the backup function and pass the sheet name to it
  createSheetBackup(sheetNameToBackup);
}

function createSheetBackup(sheetNameToBackup) {
  // Get the sheet to backup by name
  var sheetToBackup = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNameToBackup);
  if (!sheetToBackup) {
    Logger.log("Sheet with the name '" + sheetNameToBackup + "' was not found.");
    return;
  }

  // Get the backup sheet by name (Use the same name as the original sheet with "Backup" appended to it)
  var backupSheetName = sheetNameToBackup + " Backup";

  // If the backup sheet doesn't exist, create it with headers from the source sheet
  var backupSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(backupSheetName);
  if (!backupSheet) {
    backupSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(backupSheetName);
    var headers = sheetToBackup.getRange(1, 1, 1, sheetToBackup.getLastColumn()).getValues();
    backupSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);

    // If this is the first backup of the sheet, set last backed up row to 1 (to avoid copying header again)
    PropertiesService.getScriptProperties().setProperty('lastBackedUpRow', '1');
  }

  // Get the last backed up row from Script Properties
  var lastBackedUpRow = parseInt(PropertiesService.getScriptProperties().getProperty('lastBackedUpRow') || '1', 10);

  // Get the data from the sheet to backup, excluding the header row
  var lastRow = sheetToBackup.getLastRow();
  var dataRange = (lastRow > lastBackedUpRow) ? sheetToBackup.getRange(lastBackedUpRow + 1, 1, lastRow - lastBackedUpRow, sheetToBackup.getLastColumn()) : null;
  if (!dataRange) {
    // If no new data, log the message and abort backup
    Logger.log('No new data found in the sheet to backup. Backup aborted.');
    return;
  }

  var data = dataRange.getValues();

  // Copy the new data to the backup sheet
  backupSheet.getRange(backupSheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);

  // Find the last row with data in the sheet to backup
  var lastDataRow = data.length + lastBackedUpRow;

  // Update the last backed up row in Script Properties
  PropertiesService.getScriptProperties().setProperty('lastBackedUpRow', lastDataRow.toString());

  // Log the message about the successful backup
  Logger.log('Backup created successfully! Backup sheet name: ' + backupSheetName);
}





function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Run Backup', 'getSheetName')
    .addToUi();
}
