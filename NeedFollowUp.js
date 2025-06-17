/**
 * Analyzes attendance entries in 'Event Attendance' to identify
 * first-time attendees and those needing follow-up based on attendance gap,
 * using Full Name (Column B) for identification.
 */
function processEventAttendanceForFollowUpByName() { // Changed function name to indicate name-based
  Logger.log('Processing Event Attendance for follow-up by Name...');

  // --- Configuration ---
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Tab names
  const eventAttendanceTabName = 'Event Attendance';
  const sundayServiceTabName = 'Sunday Service'; // Ensure this is the exact name of the Sunday Service tab

  // Column indices (0-based) for data reading
  // UPDATED: Using Full Name (Column B) for identification
  const eventAttendanceNameCol = 1; // Column B: Full Name in Event Attendance
  const eventAttendanceDateCol = 10; // Column K: Date/Timestamp in Event Attendance - Confirmed
  // UPDATED: Using Full Name (Column B) for identification in Sunday Service
  const sundayServiceNameCol = 1; // Column B: Full Name in Sunday Service
  const sundayServiceDateCol = 4; // Column E: Date in Sunday Service (Assuming "Timestamp")

  // Column indices (0-based) for writing results in Event Attendance
  const firstTimeCol = 11; // Column L: 'First-Time' - Confirmed
  const needFollowUpCol = 12; // Column M: 'Need Follow-up?' - Confirmed

  // Header row number (1-based) - Assuming row 1 is header
  const headerRow = 1;

    // Define the threshold for follow-up gap in days
  const followUpThresholdDays = 30;


  // --- Get Sheets ---
  const eventAttendanceSheet = ss.getSheetByName(eventAttendanceTabName);
  const sundayServiceSheet = ss.getSheetByName(sundayServiceTabName);

  if (!eventAttendanceSheet) {
    Logger.log(`Error: Sheet named "${eventAttendanceTabName}" not found.`);
    SpreadsheetApp.getUi().alert(`Error: Sheet named "${eventAttendanceTabName}" not found.`);
    return;
  }
  if (!sundayServiceSheet) {
    Logger.log(`Error: Sheet named "${sundayServiceTabName}" not found.`);
    SpreadsheetApp.getUi().alert(`Error: Sheet named "${sundayServiceTabName}" not found.`);
    return;
  }

  // --- Read All Historical Attendance Data ---
  // UPDATED: Map key is standardized Full Name
  const allAttendanceDates = new Map(); // Map: Standardized Full Name -> Sorted Array of Date Objects

  // Read data from 'Sunday Service'
  const sundayServiceData = sundayServiceSheet.getDataRange().getValues();
  Logger.log(`Reading ${sundayServiceData.length} rows from "${sundayServiceTabName}".`);
  for (let i = headerRow; i < sundayServiceData.length; i++) {
    const row = sundayServiceData[i];
    const name = row[sundayServiceNameCol]; // Read name
    const date = row[sundayServiceDateCol];

    if (name && date instanceof Date) { // Process only rows with Name and valid Date
      const standardizedName = String(name).trim().toUpperCase(); // Standardize name
      if (standardizedName !== '') { // Ensure name is not just whitespace after trimming
          if (!allAttendanceDates.has(standardizedName)) {
            allAttendanceDates.set(standardizedName, []);
          }
          allAttendanceDates.get(standardizedName).push(date);
      }
    }
  }

  // Read data from 'Event Attendance' (all rows to build historical data)
  const eventAttendanceDataFull = eventAttendanceSheet.getDataRange().getValues();
  Logger.log(`Reading ${eventAttendanceDataFull.length} rows from "${eventAttendanceTabName}" for history.` );
    // Skip header row and the columns we will write to later
  const processableEventAttendanceData = eventAttendanceDataFull.slice(headerRow);


  for (let i = 0; i < processableEventAttendanceData.length; i++) {
    const row = processableEventAttendanceData[i];
    const name = row[eventAttendanceNameCol]; // Read name
    // Use the date from the correct column (K) for building historical data
    const date = row[eventAttendanceDateCol];

    if (name && date instanceof Date) { // Process only rows with Name and valid Date
        const standardizedName = String(name).trim().toUpperCase(); // Standardize name
        if (standardizedName !== '') { // Ensure name is not just whitespace after trimming
          if (!allAttendanceDates.has(standardizedName)) {
            allAttendanceDates.set(standardizedName, []);
          }
          allAttendanceDates.get(standardizedName).push(date);
        }
    }
  }

  // Sort dates for each person (by standardized name) chronologically
  allAttendanceDates.forEach(dates => {
    dates.sort((a, b) => a.getTime() - b.getTime());
  });

  Logger.log(`Built combined attendance history for ${allAttendanceDates.size} unique names.`);

  // --- Process 'Event Attendance' Entries to Populate Columns L and M ---
  const resultsToWrite = [];

  // Iterate through Event Attendance data rows (starting from the row after header)
  // We use the full data just to iterate, but calculate based on the combined history
  for (let i = headerRow; i < eventAttendanceDataFull.length; i++) {
    const row = eventAttendanceDataFull[i];
    const name = row[eventAttendanceNameCol]; // Read name for processing
    const rowNumber = i + headerRow + 1; // Calculate 1-based row number for logging
    // Use the date from the correct column (K) for the current entry being processed
    const currentDate = row[eventAttendanceDateCol];

    let firstTimeFlag = '';
    let needFollowUpFlag = '';

    // Log the start of processing for this row
    Logger.log(`--- Processing Row ${rowNumber} (Name: ${name}) ---`);

    // Process only rows with a valid Name and a valid Date in column K
    if (name && currentDate instanceof Date) {
        const standardizedName = String(name).trim().toUpperCase(); // Standardize name for lookup

        if (standardizedName !== '') { // Ensure name is not just whitespace
          // Get the sorted list of dates for this person (by standardized name)
          const personDates = allAttendanceDates.get(standardizedName) || [];
          Logger.log(`Row ${rowNumber}: Found ${personDates.length} attendance dates for name "${standardizedName}".`);
          // Optional: Log all dates found for the person (can be very verbose for many dates)
          // Logger.log(`Row ${rowNumber}: Dates found for "${standardizedName}": ${personDates.map(d => d.toISOString()).join(', ')}`);


          // --- Find the most recent attendance date *strictly before* the currentDate ---
          let lastAttendanceDate = null;
          for (const date of personDates) {
            // Use getTime() for precise comparison including time
            if (date.getTime() < currentDate.getTime()) {
              // Keep updating lastAttendanceDate as we find later dates that are still before currentDate
              lastAttendanceDate = date;
            } else {
              // Since personDates is sorted, once we hit a date >= currentDate, stop.
              break;
            }
          }

          // Log the current date and the last attendance date found
          Logger.log(`Row ${rowNumber}: Current Date (K): ${currentDate ? currentDate.toISOString() : 'Invalid Date'}`);
          Logger.log(`Row ${rowNumber}: Last Attendance Date Found (Before Current): ${lastAttendanceDate ? lastAttendanceDate.toISOString() : 'None Found'}`);


          // --- Determine 'First-Time' ---
          // If lastAttendanceDate is null, it means no date was found before currentDate in the history
          if (lastAttendanceDate === null) {
            firstTimeFlag = 'YES';
            Logger.log(`Row ${rowNumber}: Determined 'First-Time' = YES (No previous attendance found).`);
          } else {
            firstTimeFlag = ''; // Leave blank if a previous date was found
             Logger.log(`Row ${rowNumber}: Determined 'First-Time' = (Previous attendance found).`);
          }

          // --- Determine 'Need Follow-up?' ---
          // This should be based on the MOST RECENT overall attendance date for the person
          // compared to TODAY's date (when the script is run), not necessarily the current row's date.
          // Make sure 'personDates' is sorted, which it is already.
          if (personDates.length > 0) {
            const latestOverallAttendanceDate = personDates[personDates.length - 1]; // Get the very last date in the sorted array
            const today = new Date(); // Get today's date for comparison
            today.setHours(0, 0, 0, 0); // Normalize today's date to midnight for consistent day calculation

            const timeDiffSinceLastAttendance = today.getTime() - latestOverallAttendanceDate.getTime();
            const daysDiffSinceLastAttendance = timeDiffSinceLastAttendance / (1000 * 60 * 60 * 24);

            Logger.log(`Row ${rowNumber}: Latest overall attendance date: ${latestOverallAttendanceDate.toISOString()}`);
            Logger.log(`Row ${rowNumber}: Days since latest overall attendance (compared to today): ${daysDiffSinceLastAttendance}`);

            if (daysDiffSinceLastAttendance >= followUpThresholdDays) {
              needFollowUpFlag = 'YES';
              Logger.log(`Row ${rowNumber}: Determined 'Need Follow-up?' = YES (Last overall attendance > ${followUpThresholdDays} days ago).`);
            } else {
              needFollowUpFlag = '';
              Logger.log(`Row ${rowNumber}: Determined 'Need Follow-up?' = (Last overall attendance < ${followUpThresholdDays} days ago).`);
            }
          } else {
            // If no attendance found for this person at all, they don't need follow-up based on attendance gap.
            // They would be marked 'First-Time' if this is their first entry.
            needFollowUpFlag = '';
            Logger.log(`Row ${rowNumber}: Determined 'Need Follow-up?' = (No attendance history for this person).`);
          }

        } else {
          // If Name is just whitespace after trimming, skip processing for this row's logic
          firstTimeFlag = '';
          needFollowUpFlag = '';
          Logger.log(`Row ${rowNumber}: Skipping processing - Name in column B is blank or whitespace.`);
        }


    } else {
       // If Name is missing/invalid or Date is missing/invalid in column K, leave flags blank
        firstTimeFlag = '';
        needFollowUpFlag = '';
        Logger.log(`Row ${rowNumber}: Skipping processing - Missing Name in column B or invalid Date in column K.`);
    }
     Logger.log(`--- Finished Processing Row ${rowNumber} ---`);


     // Add the results for this row to our array
     // Ensure the array has 2 elements corresponding to the 2 output columns
     resultsToWrite.push([firstTimeFlag, needFollowUpFlag]);
  }

  // --- Write Results Back to Sheet ---
  if (resultsToWrite.length > 0) {
      // Determine the range to write to (starting from the row after header, columns L and M)
      // The range should be 2 columns wide (L and M)
    const targetRange = eventAttendanceSheet.getRange(headerRow + 1, firstTimeCol + 1, resultsToWrite.length, 2);
    targetRange.setValues(resultsToWrite);
    Logger.log(`Successfully wrote results for ${resultsToWrite.length} rows in "${eventAttendanceTabName}".`);
  } else {
    Logger.log(`No data rows found in "${eventAttendanceTabName}" to process.`);
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