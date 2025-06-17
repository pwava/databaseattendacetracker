function combineNamesOnFormSubmit(e) {
  // Get the active spreadsheet.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get the sheet named "Sunday Service".
  var sheet = ss.getSheetByName("Sunday Service"); 

  // If the sheet doesn't exist, log an error and exit.
  if (!sheet) {
    Logger.log("Sheet 'Sunday Service' not found.");
    return; 
  }
  
  // Determine the row to process.
  // If triggered by form submit, 'e.range' gives the row of the new entry.
  var targetRow;
  if (e && e.range) {
    targetRow = e.range.getRow();
  } else {
    // Fallback if not triggered by form submit (e.g., manual run).
    // This will process the very last row with any content.
    targetRow = sheet.getLastRow(); 
  }

  // Get the first name from column C and last name from column D of the target row.
  var firstName = sheet.getRange(targetRow, 3).getValue(); // Column C
  var lastName = sheet.getRange(targetRow, 4).getValue();  // Column D

  // Ensure both names are present before combining.
  if (firstName && lastName) {
    // Combine first name and last name.
    var fullName = firstName + " " + lastName; 

    // Set the combined full name into column B of the target row.
    sheet.getRange(targetRow, 2).setValue(fullName); // Column B
  } else {
    Logger.log("First name or last name is missing in row " + targetRow + ". Full name not combined.");
  }
}