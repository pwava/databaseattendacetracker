/**
 * Analyzes attendance entries to identify first-time attendees and those needing follow-up.
 * FINAL LOGIC: A person is "First-Time" if their name appears only ONCE in the sheet.
 */
function processEventAttendanceForFollowUpByName() {
  Logger.log('Processing Event Attendance for follow-up by Name...');

  // --- Configuration ---
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventAttendanceTabName = 'Event Attendance';
  const sundayServiceTabName = 'Sunday Service';
  
  const eventAttendanceNameCol = 1; // Col B
  const eventAttendanceDateCol = 10; // Col K
  const sundayServiceNameCol = 1;   // Col B
  const sundayServiceDateCol = 4;   // Col E
  const firstTimeCol = 11;          // Col L
  const needFollowUpCol = 12;       // Col M
  const headerRow = 1;
  const followUpThresholdDays = 30;

  // --- Get Sheets ---
  const eventAttendanceSheet = ss.getSheetByName(eventAttendanceTabName);
  const sundayServiceSheet = ss.getSheetByName(sundayServiceTabName);

  if (!eventAttendanceSheet || !sundayServiceSheet) {
    Logger.log(`Error: Could not find required sheets.`);
    return;
  }

  // --- 1. Read Data into Memory ---
  const eventAttendanceData = eventAttendanceSheet.getRange(headerRow + 1, 1, eventAttendanceSheet.getLastRow() - headerRow, eventAttendanceSheet.getLastColumn()).getValues();
  const sundayServiceData = sundayServiceSheet.getRange(headerRow + 1, 1, sundayServiceSheet.getLastRow() - headerRow, sundayServiceSheet.getLastColumn()).getValues();

  // --- 2. Count Occurrences of Each Name in 'Event Attendance' ---
  const nameCounts = new Map();
  for (const row of eventAttendanceData) {
    const name = row[eventAttendanceNameCol];
    if (name) {
      const standardizedName = String(name).trim().toUpperCase();
      if (standardizedName) {
        nameCounts.set(standardizedName, (nameCounts.get(standardizedName) || 0) + 1);
      }
    }
  }
  Logger.log(`Counted occurrences for ${nameCounts.size} unique names.`);

  // --- 3. Build Historical Data for 'Need Follow-up' Logic---
  // This logic remains unchanged
  const allAttendanceDates = new Map();
  const dataSources = [
    { data: sundayServiceData, nameCol: sundayServiceNameCol, dateCol: sundayServiceDateCol },
    { data: eventAttendanceData, nameCol: eventAttendanceNameCol, dateCol: eventAttendanceDateCol }
  ];

  for (const source of dataSources) {
    for (const row of source.data) {
      const name = row[source.nameCol];
      const date = row[source.dateCol];
      if (name && date instanceof Date) {
        const standardizedName = String(name).trim().toUpperCase();
        if (standardizedName) {
          if (!allAttendanceDates.has(standardizedName)) allAttendanceDates.set(standardizedName, []);
          allAttendanceDates.get(standardizedName).push(date);
        }
      }
    }
  }
  allAttendanceDates.forEach(dates => dates.sort((a, b) => a.getTime() - b.getTime()));
  Logger.log(`Built combined attendance history for ${allAttendanceDates.size} unique names.`);

  // --- 4. Prepare Final Results Based on the New Rules ---
  const resultsToWrite = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  for (let i = 0; i < eventAttendanceData.length; i++) {
    const row = eventAttendanceData[i];
    const name = row[eventAttendanceNameCol];
    let firstTimeFlag = '';
    let needFollowUpFlag = '';
    
    if (name) {
      const standardizedName = String(name).trim().toUpperCase();
      if (standardizedName) {
        // 'First-Time' Logic: Set to 'YES' only if the total count for this name is 1.
        if (nameCounts.get(standardizedName) === 1) {
          firstTimeFlag = 'YES';
        }

        // 'Need Follow-up' Logic: This is unchanged.
        const personDates = allAttendanceDates.get(standardizedName);
        if (personDates && personDates.length > 0) {
          const latestDate = personDates[personDates.length - 1];
          const daysDiff = (today.getTime() - latestDate.getTime()) / (1000 * 60 * 60 * 24);
          if (daysDiff >= followUpThresholdDays) {
            needFollowUpFlag = 'YES';
          }
        }
      }
    }
    resultsToWrite.push([firstTimeFlag, needFollowUpFlag]);
  }
  
  // --- 5. Write All Results Back to the Sheet ---
  if (resultsToWrite.length > 0) {
    eventAttendanceSheet.getRange(headerRow + 1, firstTimeCol + 1, resultsToWrite.length, 2).setValues(resultsToWrite);
    Logger.log(`Successfully wrote results for ${resultsToWrite.length} rows.`);
  }

  Logger.log('Event Attendance follow-up script finished.');
}

// Helper function (getDateValue - kept for completeness)
function getDateValue(value) {
  if (value instanceof Date) {
    return value;
  }
  try {
    const date = new Date(value);
    if (!isNaN(date.getTime()) && date.getFullYear() > 1900) {
        return date;
    }
  } catch (e) {
    Logger.log(`Could not parse date value: ${value}. Error: ${e}`);
  }
  return null;
}
