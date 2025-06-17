/**
 * Trigger function that automatically updates attendance stats
 * when a change occurs in the spreadsheet.
 */
function onAttendanceSheetsChange(e) {
  // Optional: You could add checks here based on the event object 'e'
  // to see what type of change occurred (e.g., edit, insert_row).
  // However, for simplicity and to ensure it runs after IMPORTRANGE updates,
  // we'll just call the update function directly.

  Logger.log("Trigger detected a change. Updating attendance stats...");

  // Call your main function to update the stats sheet
  updateAttendanceStatsSheet();

  Logger.log("Attendance stats update triggered.");
}