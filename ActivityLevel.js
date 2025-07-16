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
  
  // --- NEW: Get the date for 12 months ago ---
  const twelveMonthsAgo = new Date();
  twelveMonthsAgo.setMonth(twelveMonthsAgo.getMonth() - 12);
  
  // --- MODIFIED: Read from Column E (attendance) and H (last date) ---
  const dataRange = sheet.getRange("E2:H" + lastRow);
  const dataValues = dataRange.getValues();

  const activityLevels = dataValues.map(([count, colF, colG, lastAttended], i) => {
    
    // --- NEW: Logic to check for "Archive" status first ---
    // It checks if column H has a valid date that is older than 12 months.
    if (lastAttended && lastAttended instanceof Date && lastAttended < twelveMonthsAgo) {
      Logger.log(`Row ${i + 2}: Last attended on ${lastAttended.toDateString()} -> Archive`);
      return ["Archive"];
    }

    // If not archived, run the original logic based on attendance count
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

  // Write results to Column L (the 12th column), starting from row 2.
  sheet.getRange(2, 12, activityLevels.length, 1).setValues(activityLevels);
}
