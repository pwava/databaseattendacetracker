function updateActivityLevels() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("attendance stats");
  if (!sheet) {
    Logger.log("Sheet 'attendance stats' not found.");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("No data to process.");
    return;
  }

  // Read from Column E (attendance count), starting at row 2
  const attendanceData = sheet.getRange("E2:E" + lastRow).getValues();

  const activityLevels = attendanceData.map(([count], i) => {
    if (count === "" || isNaN(count)) {
      Logger.log(`Row ${i + 2}: Empty or invalid -> ""`);
      return [""];
    } else if (count >= 12) {
      Logger.log(`Row ${i + 2}: ${count} -> Core`);
      return ["Core"];
    } else if (count >= 3) {
      Logger.log(`Row ${i + 2}: ${count} -> Active`);
      return ["Active"];
    } else {
      Logger.log(`Row ${i + 2}: ${count} -> Inactive`);
      return ["Inactive"];
    }
  });

  // Write results to Column K (11), starting from K2
  sheet.getRange(2, 12, activityLevels.length, 1).setValues(activityLevels);
}