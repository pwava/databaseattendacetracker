/**
 * Functions related to pulling data from Sunday Service to Service Attendance
 * and managing related triggers.
 */

function pullSundayServiceToServiceAttendance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sundayServiceSheet = ss.getSheetByName("Sunday Service");
  const serviceAttendanceSheet = ss.getSheetByName("Service Attendance");

  if (!sundayServiceSheet || !serviceAttendanceSheet) {
    Logger.log("❌ Required sheets not found: 'Sunday Service' or 'Service Attendance'.");
    if (SpreadsheetApp.getUi()) { // Check if UI context exists before alerting
      SpreadsheetApp.getUi().alert("Error: Sunday Service or Service Attendance sheet not found");
    }
    return;
  }

  const sundayData = sundayServiceSheet.getDataRange().getValues();
  if (sundayData.length <= 1) {
    Logger.log("❌ No data (or only headers) in Sunday Service sheet.");
    if (SpreadsheetApp.getUi()) { // Check if UI context exists
      SpreadsheetApp.getUi().alert("No Data", "The 'Sunday Service' sheet is empty or only contains headers. No data to transfer.", SpreadsheetApp.getUi().ButtonSet.OK);
    }
    return;
  }

  const serviceData = serviceAttendanceSheet.getDataRange().getValues();
  const existingEntries = new Set();

  // Start from row 1 (index 1) if row 0 is headers in serviceData
  for (let i = 1; i < serviceData.length; i++) {
    if (serviceData[i][1] && serviceData[i][4]) { // Full Name and Timestamp
      try {
        const key = `${String(serviceData[i][1]).trim().toLowerCase()}_${new Date(serviceData[i][4]).getTime()}`;
        existingEntries.add(key);
      } catch (e) {
        Logger.log(`Error processing existing entry key at serviceData row ${i+1}: ${e.toString()}`);
      }
    }
  }

  const newEntries = [];
  let skippedCount = 0;

  // Start from row 1 (index 1) if row 0 is headers in sundayData
  for (let i = 1; i < sundayData.length; i++) {
    const row = sundayData[i];

    const personalId = row[0];
    const fullName = row[1];
    const firstName = row[2];
    const lastName = row[3];
    const timestamp = row[4];
    const firstTime = row[5];
    const email = row[6];

    if (!personalId || !fullName || !timestamp) {
      let missing = [];
      if (!personalId) missing.push("Personal ID (Column A)");
      if (!fullName) missing.push("Full Name (Column B)");
      if (!timestamp) missing.push("Timestamp (Column E)");
      Logger.log(`⚠️ Skipping row ${i + 1} from 'Sunday Service': Missing ${missing.join(', ')}.`);
      continue;
    }

    try {
      const entryKey = `${String(fullName).trim().toLowerCase()}_${new Date(timestamp).getTime()}`;
      if (existingEntries.has(entryKey)) {
        skippedCount++;
        continue;
      }

      newEntries.push([
        personalId || "",
        fullName,
        firstName || "",
        lastName || "",
        timestamp,
        firstTime || "No",
        email || "",
        "", // Assuming an empty 8th column (e.g., for Notes)
        timestamp // Assuming 9th column is also a timestamp (e.g., Date Added)
      ]);
    } catch (e) {
        Logger.log(`Error processing new entry key at sundayData row ${i+1} for '${fullName}': ${e.toString()}`);
        continue; // Skip this problematic row
    }
  }

  if (newEntries.length > 0) {
    const lastRow = serviceAttendanceSheet.getLastRow();
    // Determine startRow correctly: if sheet is completely empty (lastRow=0), start at 1.
    // If sheet has content, start at lastRow + 1.
    const startRow = (lastRow === 0 && serviceAttendanceSheet.getRange("A1").isBlank()) ? 1 : lastRow + 1;

    serviceAttendanceSheet.getRange(startRow, 1, newEntries.length, 9) // Assuming 9 columns for newEntries
      .setValues(newEntries);

    Logger.log(`✅ Successfully added ${newEntries.length} new entries to 'Service Attendance'.`);
    Logger.log(`ℹ️ Skipped ${skippedCount} duplicate entries.`);

    if (SpreadsheetApp.getUi()) { // Check if UI context exists
      SpreadsheetApp.getUi().alert(
        'Transfer Complete!',
        `Successfully transferred ${newEntries.length} new entries to 'Service Attendance'.\n` +
        `Skipped ${skippedCount} duplicate entries.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } else {
    Logger.log("ℹ️ No new entries to transfer after checking for duplicates and valid data.");
    if (SpreadsheetApp.getUi()) { // Check if UI context exists
      SpreadsheetApp.getUi().alert(
        'No New Entries',
        'All valid entries from the \'Sunday Service\' sheet are already in \'Service Attendance\', or no new valid entries were found.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  }
}

function setupFormSubmitTrigger() {
  // This function is typically run manually, so SpreadsheetApp.getUi() is fine here.
  const ui = SpreadsheetApp.getUi(); 

  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onFormSubmitTransfer') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log("Removed existing 'onFormSubmitTransfer' trigger.");
    }
  });

  ScriptApp.newTrigger('onFormSubmitTransfer')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onFormSubmit()
    .create();

  Logger.log("✅ Form submit trigger created for 'onFormSubmitTransfer'.");
  ui.alert(
    'Trigger Created!',
    'Automatic transfer has been set up. New form submissions to linked forms targeting "Sunday Service" will be processed.',
    ui.ButtonSet.OK
  );
}

function onFormSubmitTransfer(e) {
    Utilities.sleep(5000); 

    if (!e || !e.range) { // Check if e and e.range exist
        Logger.log("Error: Event object 'e' or 'e.range' was undefined. This function might have been run manually or the event data was incomplete.");
        return;
    }
    const range = e.range;
    const row = range.getRow();
    const sheet = range.getSheet();

    if (sheet.getName() !== "Sunday Service") {
        Logger.log(`Skipping onFormSubmitTransfer for sheet '${sheet.getName()}' as it's not 'Sunday Service'. Event was for row ${row}.`);
        return;
    }

    const numberOfColumnsToRead = 7; 
    let initialRowData = sheet.getRange(row, 1, 1, numberOfColumnsToRead).getValues()[0];
    let fullName = initialRowData[1]; 
    let timestamp = initialRowData[4]; 

    if (!fullName || !timestamp) {
        Logger.log(`❌ CRITICAL ERROR: Form submission on row ${row} in sheet '${sheet.getName()}' appears to be missing essential form data: Full Name ('${fullName}') or Timestamp ('${timestamp}'). Skipping.`);
        return;
    }

    let personalId = initialRowData[0]; 
    const MAX_RETRIES_FOR_ID = 50;
    const RETRY_DELAY_MS_FOR_ID = 2500;

    if (!personalId) {
        Logger.log(`ℹ️ Personal ID missing for '${fullName}' (Row ${row} on '${sheet.getName()}'). Waiting for it to be generated...`);
        for (let i = 0; i < MAX_RETRIES_FOR_ID; i++) {
            Utilities.sleep(RETRY_DELAY_MS_FOR_ID);
            personalId = sheet.getRange(row, 1).getValue(); 
            if (personalId) {
                Logger.log(`✅ Personal ID for '${fullName}' (Row ${row}) populated after ${i + 1} attempt(s). ID: ${personalId}`);
                break; 
            }
            Logger.log(`⏳ Attempt ${i + 1} of ${MAX_RETRIES_FOR_ID}: Personal ID still missing for '${fullName}' (Row ${row}).`);
        }
    }

    if (!personalId) {
        Logger.log(`❌ SKIPPING ROW ${row} ('${fullName}' from '${sheet.getName()}'): Personal ID (Column A) was NOT generated/found after ${MAX_RETRIES_FOR_ID} retries.`);
        return;
    }

    const finalRowData = sheet.getRange(row, 1, 1, numberOfColumnsToRead).getValues()[0];
    personalId = finalRowData[0];
    fullName = finalRowData[1];
    const firstName = finalRowData[2];
    const lastName = finalRowData[3];
    timestamp = finalRowData[4];
    const firstTime = finalRowData[5];
    const email = finalRowData[6];

    if (!personalId || !fullName || !timestamp) { // Safeguard check
        let missing = [];
        if (!personalId) missing.push("Personal ID (Column A - safeguard check)");
        if (!fullName) missing.push("Full Name (Column B - safeguard check)");
        if (!timestamp) missing.push("Timestamp (Column E - safeguard check)");
        Logger.log(`ℹ️ Safeguard check failed for row ${row} in sheet '${sheet.getName()}': Missing ${missing.join(', ')}.`);
        return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet(); // Already defined, but fine to redefine scope
    const serviceAttendanceSheet = ss.getSheetByName("Service Attendance");

    if (!serviceAttendanceSheet) {
        Logger.log("❌ 'Service Attendance' sheet not found for onFormSubmitTransfer.");
        return;
    }

    const serviceData = serviceAttendanceSheet.getDataRange().getValues();
    const existingEntries = new Set();
    for (let i = 1; i < serviceData.length; i++) { 
        const existingFullName = serviceData[i][1];
        const existingTimestamp = serviceData[i][4];
        if (existingFullName && existingTimestamp) {
            try {
              const key = `${String(existingFullName).trim().toLowerCase()}_${new Date(existingTimestamp).getTime()}`;
              existingEntries.add(key);
            } catch (ex) {
              Logger.log(`Error processing existing entry key during onFormSubmitTransfer at serviceData row ${i+1}: ${ex.toString()}`);
            }
        }
    }
    let entryKey;
    try {
      entryKey = `${String(fullName).trim().toLowerCase()}_${new Date(timestamp).getTime()}`;
    } catch (ex) {
      Logger.log(`Error creating entryKey for '${fullName}' (Row ${row}) during onFormSubmitTransfer: ${ex.toString()}. Skipping.`);
      return;
    }

    if (existingEntries.has(entryKey)) {
        Logger.log(`ℹ️ Form submission for '${fullName}' (Row ${row}) skipped: Duplicate entry found in 'Service Attendance'.`);
        return;
    }

    const newEntry = [
        personalId || "",
        fullName,
        firstName || "",
        lastName || "",
        timestamp,
        firstTime || "No",
        email || "",
        "", 
        timestamp 
    ];

    const lastRowSA = serviceAttendanceSheet.getLastRow();
    const targetRow = (lastRowSA === 0 && serviceAttendanceSheet.getRange("A1").isBlank()) ? 1 : lastRowSA + 1;

    serviceAttendanceSheet.getRange(targetRow, 1, 1, newEntry.length).setValues([newEntry]);
    Logger.log(`✅ Auto-transferred new form submission for '${fullName}' (ID: ${personalId}) from '${sheet.getName()}' to 'Service Attendance' sheet, row ${targetRow}.`);
}

/**
 * Adds the "Data Transfer" menu.
 * This function will be called by the master onOpen(e) function in another script file
 * within the same project.
 */
function addTransferMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Data Transfer')
    .addItem('📥 Pull Sunday Service → Service Attendance', 'pullSundayServiceToServiceAttendance')
    // The 'Set Up Automatic Transfer' item has been removed from here
    .addToUi();
  Logger.log("✅ Data Transfer menu definition attempted by addTransferMenu. 'Set Up Automatic Transfer' item removed.");
}

// Removed the onOpen() function from this file.
// It will be handled by a single onOpen() in the project,
// which will call addTransferMenu().