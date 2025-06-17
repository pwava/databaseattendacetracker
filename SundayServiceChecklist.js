/**
 * Sunday Registration & Service Utilities System (Primary Project)
 * Manages Sunday Service registration, attendance, and Service Stats.
 * Contains all shared helper functions for use by other Apps Script projects
 * bound to the same spreadsheet.
 */

// Define the names of sheets that might contain person IDs locally within *this* spreadsheet.
// This list is used for ID generation logic to ensure uniqueness.
const LOCAL_ID_SHEETS = ["Sunday Registration", "Service Attendance", "Sunday Service", "Event Registration", "Event Attendance"];

// --- Sunday Registration Functions ---

/**
 * Creates or recreates the main "Sunday Registration" sheet.
 * Prompts the user if the sheet already exists.
 */
function createSundayRegistrationSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let regSheet = ss.getSheetByName("Sunday Registration");
  if (regSheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Sheet Already Exists',
      'Sunday Registration sheet already exists. Do you want to recreate it? This will clear all existing data.',
      ui.ButtonSet.YES_NO
    );
    if (response === ui.Button.YES) {
      ss.deleteSheet(regSheet);
    } else {
      return; // User chose not to recreate
    }
  }
  regSheet = ss.insertSheet("Sunday Registration");
  setupSundayRegistrationSheetLayout(regSheet);
  populateSundayRegistrationList(regSheet);
  Logger.log("✅ Sunday Registration sheet created successfully! Person IDs are populated via new logic.");
  SpreadsheetApp.getUi().alert(
    'Registration Sheet Created!',
    'Sunday Registration sheet has been created and populated with active members.\n\n' +
    'Person IDs in Column A are fetched/generated based on Directory, local sheets, or new.\n\n' +
    'The registration team can now:\n' +
    '1. Enter the service date in cell B2\n' +
    '2. Check the boxes for attendees\n' +
    '3. Click "Submit Attendance" from the "📋 Sunday Check-in" menu to transfer to Service Attendance sheet\n\n' +
    'Menus have been added/updated for easy access to functions.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Sets up the initial layout, headers, and basic formatting for the Sunday registration sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to format.
 */
function setupSundayRegistrationSheetLayout(sheet) {
  sheet.clear();
  sheet.getRange("A1").setValue("🏛️ SUNDAY SERVICE REGISTRATION").setFontSize(16).setFontWeight("bold");
  sheet.getRange("A1:E1").merge().setHorizontalAlignment("center");
  sheet.getRange("A2").setValue("📅 Service Date:");
  sheet.getRange("B2").setValue(new Date()).setNumberFormat("MM/dd/yyyy");
  sheet.getRange("A3").setValue("📝 Instructions: Check the box next to each person who is present today");
  sheet.getRange("A3:E3").merge();
  sheet.getRange("A4").setValue("🔄 Refresh List");
  sheet.getRange("B4").setValue("✅ Submit Attendance");
  sheet.getRange("C4").setValue("🧹 Clear All Checks");
  sheet.getRange("D4").setValue("Status: Ready");

  const headers = ["ID", "Full Name", "First Name", "Last Name", "✓ Present"];
  sheet.getRange("A5:E5").setValues([headers]).setFontWeight("bold").setBackground("#4285f4").setFontColor("white");

  sheet.setColumnWidth(1, 70);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 140);
  sheet.setColumnWidth(5, 80);
  sheet.hideColumns(1);

  sheet.getRange("A1:E4").setBackground("#f8f9fa");
  sheet.getRange("A2:B2").setBackground("#e3f2fd");
  sheet.getRange("A4:D4").setBackground("#fff3e0");
  sheet.setFrozenRows(5);
  Logger.log("✅ Sunday Registration sheet layout created.");
}

/**
 * Populates the "Sunday Registration" list SOLELY from the Directory.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} regSheet The registration sheet to populate (optional, defaults to active sheet).
 */
function populateSundayRegistrationList(regSheet = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!regSheet) {
    regSheet = ss.getSheetByName("Sunday Registration");
    if (!regSheet) { Logger.log("❌ Sunday Registration sheet not found for populateSundayRegistrationList"); return; }
  }

  const directoryMap = getDirectoryDataMap();
  if (directoryMap.size === 0) {
    SpreadsheetApp.getUi().alert("Warning", "The Directory is empty or could not be loaded. Please ensure the Directory Spreadsheet URL is set correctly and the 'Directory' sheet contains data.", SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log("⚠️ Directory map is empty or failed to load for populateSundayRegistrationList. Cannot populate list.");
    return;
  }

  const serviceAttendanceIdMap = getLocalSheetIdMap("Service Attendance", 1, 2);

  let nextGeneratedId = findHighestIdInDirectory();
  if (nextGeneratedId === 0) {
    nextGeneratedId = findHighestIdInLocalSheets(LOCAL_ID_SHEETS);
  }
  Logger.log(`Initial base for nextGeneratedId (starting with Directory, then local): ${nextGeneratedId}`);

  const personsForRegistration = [];
  const processedNewPersonsInThisRun = new Map();

  for (const [normalizedFullName, directoryEntry] of directoryMap.entries()) {
    let personId = directoryEntry.id;
    let firstName = directoryEntry.firstName;
    let lastName = directoryEntry.lastName;
    const fullName = directoryEntry.originalFullName;

    if (!personId) {
      const serviceEntryId = serviceAttendanceIdMap.get(normalizedFullName);
      const alreadyProcessedNew = processedNewPersonsInThisRun.get(normalizedFullName);

      if (serviceEntryId) {
        personId = serviceEntryId;
      } else if (alreadyProcessedNew) {
        personId = alreadyProcessedNew.id;
        firstName = alreadyProcessedNew.firstName || firstName;
        lastName = alreadyProcessedNew.lastName || lastName;
      } else {
        nextGeneratedId++;
        personId = String(nextGeneratedId);
        processedNewPersonsInThisRun.set(normalizedFullName, { id: personId, firstName: firstName, lastName: lastName });
        Logger.log(`Generated new ID ${personId} for ${fullName} (from Directory but ID missing for Sunday Registration).`);
      }
    }

    if (!firstName && !lastName && fullName) {
      const nameParts = fullName.split(/\s+/);
      firstName = nameParts[0] || "";
      lastName = nameParts.length > 1 ? nameParts.slice(1).join(" ") : "";
    }
    personsForRegistration.push([personId, fullName, firstName, lastName, false]);
  }


  personsForRegistration.sort((a, b) => (String(a[3]) || "").toLowerCase().localeCompare((String(b[3]) || "").toLowerCase()));

  const lastDataRowOnSheet = regSheet.getLastRow();
  if (lastDataRowOnSheet > 5) {
    regSheet.getRange(6, 1, lastDataRowOnSheet - 5, 5).clearContent().clearFormat();
  }
  if (personsForRegistration.length > 0) {
    const startRow = 6;
    regSheet.getRange(startRow, 1, personsForRegistration.length, 5).setValues(personsForRegistration);
    const checkboxRange = regSheet.getRange(startRow, 5, personsForRegistration.length, 1);
    checkboxRange.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
    regSheet.getRange(startRow, 1, personsForRegistration.length, 5).setBorder(true, true, true, true, true, true);
    refreshRowFormatting(regSheet, startRow, personsForRegistration.length);
  }
  regSheet.getRange("D4").setValue(`Status: ${personsForRegistration.length} members loaded`);
  Logger.log(`✅ Sunday Registration list populated with ${personsForRegistration.length} members. IDs fetched/generated.`);
}

/**
 * Adds a new person to the Sunday Registration sheet (quick add).
 * This function now uses the centralized ID resolution logic.
 */
function addPersonToSundayRegistration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Sunday Registration");
  const ui = SpreadsheetApp.getUi();

  if (!regSheet) {
    ui.alert("Error", "Sunday Registration sheet not found", ui.ButtonSet.OK);
    return;
  }

  const nameResponse = ui.prompt('Add Person', 'Enter the full name:', ui.ButtonSet.OK_CANCEL);
  if (nameResponse.getSelectedButton() !== ui.Button.OK) return;
  const fullNameEntered = String(nameResponse.getResponseText() || "").trim();
  if (!fullNameEntered) {
    ui.alert('Input Error', 'Please enter a valid name', ui.ButtonSet.OK);
    return;
  }

  // Check for duplicate on the CURRENT Sunday registration sheet first
  const lastDataRow = regSheet.getLastRow();
  if (lastDataRow >= 6) {
    const existingNamesOnCurrentSheet = regSheet.getRange(6, 2, lastDataRow - 5, 1).getValues();
    if (existingNamesOnCurrentSheet.some(row => row[0] && String(row[0]).trim().toLowerCase() === fullNameEntered.toLowerCase())) {
      ui.alert('Duplicate Entry', 'This person is already in the current Sunday registration list. No need to add again.', ui.ButtonSet.OK);
      return;
    }
  }

  // --- NEW: Use the centralized ID resolver ---
  const personDetails = resolvePersonIdAndDetails(fullNameEntered); // Pass only name; resolver fetches all sheets
  const personIdToAdd = personDetails.id;
  const firstNameToAdd = personDetails.firstName;
  const lastNameToAdd = personDetails.lastName;
  // --- END NEW ---

  const nextSheetRow = (lastDataRow < 5) ? 6 : lastDataRow + 1;
  const newRowData = [personIdToAdd, fullNameEntered, firstNameToAdd, lastNameToAdd, false];
  regSheet.getRange(nextSheetRow, 1, 1, 5).setValues([newRowData]);
  regSheet.getRange(nextSheetRow, 5).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  const newRowRange = regSheet.getRange(nextSheetRow, 1, 1, 5);
  newRowRange.setBorder(true, true, true, true, true, true);
  refreshRowFormatting(regSheet);

  ui.alert('Person Added!', `${fullNameEntered} has been added with ID ${personIdToAdd}.`, ui.ButtonSet.OK);
  Logger.log(`✅ Manually added ${fullNameEntered} (ID: ${personIdToAdd}) to Sunday registration list.`);
}


/**
 * Submits checked-in attendees from "Sunday Registration" to the "Service Attendance" sheet.
 */
function submitSundayRegistrationAttendance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Sunday Registration");
  if (!regSheet) {
    SpreadsheetApp.getUi().alert("Error", "Sunday Registration sheet not found", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const serviceDateValue = regSheet.getRange("B2").getValue();
  if (!serviceDateValue || !(serviceDateValue instanceof Date) || isNaN(serviceDateValue.getTime())) {
    SpreadsheetApp.getUi().alert("Input Error", "Please enter a valid service date in cell B2.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const spreadsheetTimezone = ss.getSpreadsheetTimeZone();
  const formattedServiceDate = Utilities.formatDate(serviceDateValue, spreadsheetTimezone, "MM/dd/yyyy");

  regSheet.getRange("D4").setValue("Status: Processing...");

  try {
    const serviceSheet = ss.getSheetByName("Service Attendance");
    if (!serviceSheet) {
      SpreadsheetApp.getUi().alert("Error", "'Service Attendance' sheet not found. Please create it manually or ensure it exists.", SpreadsheetApp.getUi().ButtonSet.OK);
      regSheet.getRange("D4").setValue("Status: Error - Service Attendance sheet missing");
      throw new Error("'Service Attendance' sheet not found");
    }

    const directoryMap = getDirectoryDataMap();

    const lastRegDataRow = regSheet.getLastRow();
    if (lastRegDataRow < 6) {
      regSheet.getRange("D4").setValue("Status: No members to process");
      SpreadsheetApp.getUi().alert("No Members", "No members listed to process for attendance.", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const regData = regSheet.getRange(6, 1, lastRegDataRow - 5, 5).getValues();
    const attendanceEntries = [];
    let checkedCount = 0;

    for (const row of regData) {
      const [personId, fullName, firstName, lastName, isChecked] = row;

      if (isChecked === true && fullName && String(fullName).trim() !== "") {
        let email = "";
        const normalizedFullName = String(fullName).trim().toUpperCase();
        const directoryEntry = directoryMap.get(normalizedFullName);
        if (directoryEntry && directoryEntry.email) {
          email = directoryEntry.email;
        } else {
          Logger.log(`Email not found in Directory for ${fullName} (ID: ${personId}). Will submit blank email for Sunday Service.`);
        }

        const notes = ""; // Original script had notes column
        attendanceEntries.push([
          personId,
          fullName,
          firstName || "",
          lastName || "",
          formattedServiceDate,
          "No", // Placeholder for "Visitor" or similar column in old script
          email,
          notes,
          new Date() // Timestamp
        ]);
        checkedCount++;
      }
    }

    if (attendanceEntries.length === 0) {
      regSheet.getRange("D4").setValue("Status: No members checked");
      SpreadsheetApp.getUi().alert("No Checks", "No members were checked for attendance.", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // Ensure 'Service Attendance' sheet has headers if new/empty
    if (serviceSheet.getLastRow() < 1) {
      const serviceHeaders = ["Person ID", "Full Name", "First Name", "Last Name", "Service Date", "Is Visitor?", "Email", "Notes", "Timestamp"];
      serviceSheet.getRange(1, 1, 1, serviceHeaders.length).setValues([serviceHeaders]).setFontWeight("bold").setBackground("#4285f4").setFontColor("white");
    }

    const nextRowServiceSheet = findLastRowWithData(serviceSheet) + 1;
    serviceSheet.getRange(nextRowServiceSheet, 1, attendanceEntries.length, 9).setValues(attendanceEntries);
    // Ensure the date column in 'Service Attendance' (Col E) is formatted as a date
    serviceSheet.getRange(nextRowServiceSheet, 5, attendanceEntries.length, 1).setNumberFormat("MM/dd/yyyy");
    // Ensure the timestamp column (Col I) is formatted as Date Name
    serviceSheet.getRange(nextRowServiceSheet, 9, 1, 1).setNumberFormat("MM/dd/yyyy HH:mm:ss");


    regSheet.getRange(6, 5, lastRegDataRow - 5, 1).setValue(false); // Clear checkboxes
    regSheet.getRange("D4").setValue(`Status: ${checkedCount} attendees submitted`);
    SpreadsheetApp.getUi().alert(
      'Attendance Submitted!',
      `Successfully submitted attendance for ${checkedCount} members to 'Service Attendance' sheet.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    Logger.log(`✅ Successfully submitted ${checkedCount} Sunday Service attendance entries.`);

    // --- NEW: Call to update Service Stats after a manual Sunday registration submission ---
    populateServiceStatsSheet();

  } catch (error) {
    regSheet.getRange("D4").setValue("Status: Error occurred");
    Logger.log(`❌ Error submitting Sunday attendance: ${error.message}\n${error.stack || ""}`);
    SpreadsheetApp.getUi().alert("Error", `Error submitting Sunday attendance: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Clears all checkboxes in the "Sunday Registration" sheet.
 */
function clearAllSundayChecks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Sunday Registration");
  if (!regSheet) {
    SpreadsheetApp.getUi().alert("Error", "Sunday Registration sheet not found", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  const lastDataRow = regSheet.getLastRow();
  if (lastDataRow >= 6) {
    regSheet.getRange(6, 5, lastDataRow - 5, 1).setValue(false);
    regSheet.getRange("D4").setValue("Status: All checks cleared");
    Logger.log("✅ All Sunday checkboxes cleared");
  } else {
    regSheet.getRange("D4").setValue("Status: No checks to clear");
    Logger.log("ℹ️ No data rows found to clear Sunday checks from.");
  }
}

/**
 * Adds or re-applies checkboxes to the 'Present' column (Column E)
 * of the "Sunday Registration" sheet.
 */
function addCheckboxesToSundayRegistration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Sunday Registration");
  if (!regSheet) {
    SpreadsheetApp.getUi().alert("Error", "Sunday Registration sheet not found", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  const lastRow = regSheet.getLastRow();
  if (lastRow < 6) {
    SpreadsheetApp.getUi().alert("No Data", "No data found below row 5 to add checkboxes to. Please add member data starting row 6.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const dataRangeForNames = regSheet.getRange(6, 2, lastRow - 5, 1);
  const nameValues = dataRangeForNames.getValues();
  let rowsWithActualNames = 0;
  for (let i = 0; i < nameValues.length; i++) {
    if (String(nameValues[i][0] || "").trim() !== "") {
      rowsWithActualNames = i + 1;
    }
  }
  if (rowsWithActualNames === 0) {
    SpreadsheetApp.getUi().alert("No Names Found", "No names found in Column B (Full Name) from row 6 downwards. Cannot add checkboxes meaningfully.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  try {
    const checkboxRange = regSheet.getRange(6, 5, rowsWithActualNames, 1);
    checkboxRange.clearContent().setValue(false).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());

    const dataFormattingRange = regSheet.getRange(6, 1, rowsWithActualNames, 5);
    dataFormattingRange.setBorder(true, true, true, true, true, true);
    refreshRowFormatting(regSheet, 6, rowsWithActualNames);

    regSheet.getRange("D4").setValue(`Status: ${rowsWithActualNames} members ready`);
    SpreadsheetApp.getUi().alert('Checkboxes Added/Reformatted!', `Successfully added/reformatted checkboxes for ${rowsWithActualNames} member rows.\n\nSheet is ready:\n1. Enter service date in B2\n2. Check attendance\n3. Click Submit Attendance`, SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log(`✅ Added/Reformatted checkboxes to ${rowsWithActualNames} rows for Sunday Registration`);
  } catch (error) {
    Logger.log(`❌ Error adding/reformatting checkboxes for Sunday Registration: ${error.message}`);
    SpreadsheetApp.getUi().alert("Error", `Error adding/reformatting checkboxes for Sunday Registration: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Creates an empty "Sunday Registration" sheet.
 */
function createEmptySundayRegistrationSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let regSheet = ss.getSheetByName("Sunday Registration");
  if (regSheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert('Sheet Already Exists', 'Sunday Registration sheet already exists. Recreate it as empty?', ui.ButtonSet.YES_NO);
    if (response === ui.Button.YES) ss.deleteSheet(regSheet);
    else return;
  }
  regSheet = ss.insertSheet("Sunday Registration");
  setupSundayRegistrationSheetLayout(regSheet);
  Logger.log("✅ Empty Sunday Registration sheet created.");
  SpreadsheetApp.getUi().alert(
    'Empty Registration Sheet Created!',
    'Sunday Registration sheet is ready for manual data entry (Columns A-E for data, starting row 6).\n\n' +
    'Person IDs (Col A) will be fetched from Directory or generated if you use Refresh/Add Attendee.\n\n' +
    'INSTRUCTIONS:\n' +
    '1. Paste directory data starting row 6 (Full Name in Col B, First in C, Last in D - ID will be handled by other functions)\n' +
    '2. Use "📋 Sunday Check-in" → "🔲 Add/Reformat Checkboxes" to set up column E.\n' +
    '3. Enter service date in B2 and start checking attendance!',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Removes a person from the "Sunday Registration" list by full name.
 */
function removePersonFromSundayRegistration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Sunday Registration");
  if (!regSheet) {
    SpreadsheetApp.getUi().alert("Error", "Sunday Registration sheet not found", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  const ui = SpreadsheetApp.getUi();
  const lastDataRow = regSheet.getLastRow();
  if (lastDataRow < 6) {
    ui.alert('No People', 'No people in the list to remove (list is empty below row 5).', ui.ButtonSet.OK);
    return;
  }

  const nameResponse = ui.prompt('Remove Person', 'Enter the FULL NAME of the person to remove (case-insensitive):', ui.ButtonSet.OK_CANCEL);
  if (nameResponse.getSelectedButton() !== ui.Button.OK) return;
  const nameToRemove = String(nameResponse.getResponseText() || "").trim().toLowerCase();
  if (!nameToRemove) {
    ui.alert("No Name Entered", "No name entered to remove.", ui.ButtonSet.OK);
    return;
  }

  const allData = regSheet.getRange(6, 1, lastDataRow - 5, 5).getValues();
  let rowToDeleteInSheet = -1;
  for (let i = 0; i < allData.length; i++) {
    if (allData[i][1] && String(allData[i][1]).trim().toLowerCase() === nameToRemove) {
      rowToDeleteInSheet = i + 6;
      break;
    }
  }

  if (rowToDeleteInSheet > 0) {
    regSheet.deleteRow(rowToDeleteInSheet);
    ui.alert('Person Removed!', `'${nameResponse.getResponseText().trim()}' has been removed.`, ui.ButtonSet.OK);
    Logger.log(`✅ Removed '${nameResponse.getResponseText().trim()}' from Sunday registration list, row ${rowToDeleteInSheet}`);
    refreshRowFormatting(regSheet);
  } else {
    ui.alert('Not Found', `Person '${nameResponse.getResponseText().trim()}' not found in the Sunday registration list.`, ui.ButtonSet.OK);
  }
}

/**
 * Sorts the data in the "Sunday Registration" sheet by Last Name (Column D).
 */
function sortSundayRegistrationByLastName() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Sunday Registration");
  if (!regSheet) {
    SpreadsheetApp.getUi().alert("Error", "Sunday Registration sheet not found.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const lastDataRow = findLastRowWithData(regSheet);
  if (lastDataRow < 6) {
    SpreadsheetApp.getUi().alert("No Data", "No data to sort (list is empty below row 5).", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const numDataRows = lastDataRow - 5;
  if (numDataRows <= 0) {
    SpreadsheetApp.getUi().alert("No Data", "No data rows to sort.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const dataRange = regSheet.getRange(6, 1, numDataRows, 5);
  dataRange.sort({ column: 4, ascending: true });
  refreshRowFormatting(regSheet, 6, numDataRows);
  SpreadsheetApp.getUi().alert("Sunday list sorted by Last Name.");
  Logger.log("✅ Sunday Registration list sorted by last name.");
}

/**
 * Adds a custom menu for Sunday Registration functions.
 */
function addSundayRegistrationMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📋 Sunday Check-in')
    .addItem('📁 Get Names from Directory', 'populateSundayRegistrationList')
    .addItem('✅ Submit Attendance', 'submitSundayRegistrationAttendance')
    .addSeparator()
    .addItem('➕ Add Attendee (Quick Add)', 'addPersonToSundayRegistration')
    .addItem('🔲 Add/Reformat Checkboxes', 'addCheckboxesToSundayRegistration')
    .addItem('Sort by Last Name', 'sortSundayRegistrationByLastName')
    .addSeparator()
    .addItem('🆕 Create Empty Registration Sheet', 'createEmptySundayRegistrationSheet')
    .addItem('📊 Generate Service Stats Report', 'createServiceStatsSheet')
    .addToUi();
  Logger.log("✅ Sunday Check-in menu definition attempted by addSundayRegistrationMenu.");
}

// --- Google Form Submission Handler (part of Sunday project) ---

/**
 * Processes a new form submission from the "Sunday Service" sheet
 * and appends the data to the "Service Attendance" sheet.
 * This function is intended to be run by an 'onFormSubmit' trigger.
 * @param {GoogleAppsScript.Events.SheetsOnFormSubmit} e The form submission event object.
 */
function processSundayFormResponse(e) {
  Logger.log("Processing Sunday form response...");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const serviceFormSheet = e.range.getSheet(); // The sheet where the form response landed (Sunday Service)
  const serviceSheet = ss.getSheetByName("Service Attendance"); // The target attendance sheet

  if (!serviceSheet) {
    Logger.log("❌ Target 'Service Attendance' sheet not found for form submission processing.");
    return;
  }
  if (serviceFormSheet.getName() !== "Sunday Service") {
    Logger.log("Skipping form response: Not from 'Sunday Service' sheet.");
    return; // Only process forms from "Sunday Service"
  }

  const newRow = e.range.getValues()[0]; // The submitted row
  const headers = serviceFormSheet.getRange(1, 1, 1, serviceFormSheet.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toLowerCase());

  // Map form response columns to expected data points
  // Adjust these column indices based on your actual Google Form output headers
  const TIMESTAMP_COL_FORM_IDX = headers.indexOf("timestamp");
  const FULL_NAME_COL_FORM_IDX = headers.indexOf("full name");
  const FIRST_NAME_COL_FORM_IDX = headers.indexOf("first name");
  const LAST_NAME_COL_FORM_IDX = headers.indexOf("last name");
  const EMAIL_COL_FORM_IDX = headers.indexOf("email");
  // Add other relevant columns from your form if needed

  let timestamp = newRow[TIMESTAMP_COL_FORM_IDX];
  let fullName = String(newRow[FULL_NAME_COL_FORM_IDX] || "").trim();
  let firstName = String(newRow[FIRST_NAME_COL_FORM_IDX] || "").trim();
  let lastName = String(newRow[LAST_NAME_COL_FORM_IDX] || "").trim();
  let email = String(newRow[EMAIL_COL_FORM_IDX] || "").trim();
  const serviceDate = Utilities.formatDate(timestamp, ss.getSpreadsheetTimeZone(), "MM/dd/yyyy"); // Use form timestamp as service date

  // If first/last names aren't directly from form, try to derive from full name
  if (!firstName && !lastName && fullName) {
    const nameParts = fullName.split(/\s+/);
    firstName = nameParts[0] || "";
    lastName = nameParts.length > 1 ? nameParts.slice(1).join(" ") : "";
  }

  // --- NEW: Use the centralized ID resolver for form submissions ---
  const personDetails = resolvePersonIdAndDetails(fullName); // Resolver fetches all sheets
  const personId = personDetails.id;
  firstName = personDetails.firstName || firstName; // Prefer resolved details, fallback to form data
  lastName = personDetails.lastName || lastName;
  email = personDetails.email || email;
  // --- END NEW ---
  
  // Prepare entry for "Service Attendance" sheet
  // Headers: ["Person ID", "Full Name", "First Name", "Last Name", "Service Date", "Is Visitor?", "Email", "Notes", "Timestamp"]
  const entryToServiceAttendance = [
    personId,
    fullName,
    firstName,
    lastName,
    serviceDate, // Use formatted service date
    "No", // Assuming form submissions are not visitors, adjust if needed
    email,
    "",   // Notes (blank for form submission)
    new Date() // Timestamp of when the script processes this form response
  ];

  // Ensure 'Service Attendance' sheet has headers if new/empty
  if (serviceSheet.getLastRow() < 1) {
    const serviceHeaders = ["Person ID", "Full Name", "First Name", "Last Name", "Service Date", "Is Visitor?", "Email", "Notes", "Timestamp"];
    serviceSheet.getRange(1, 1, 1, serviceHeaders.length).setValues([serviceHeaders]).setFontWeight("bold").setBackground("#4285f4").setFontColor("white");
  }

  const nextRowServiceSheet = findLastRowWithData(serviceSheet) + 1;
  serviceSheet.getRange(nextRowServiceSheet, 1, 1, entryToServiceAttendance.length).setValues([entryToServiceAttendance]);
  serviceSheet.getRange(nextRowServiceSheet, 5, 1, 1).setNumberFormat("MM/dd/yyyy"); // Format Service Date
  serviceSheet.getRange(nextRowServiceSheet, 9, 1, 1).setNumberFormat("MM/dd/yyyy HH:mm:ss"); // Format Timestamp

  Logger.log(`✅ Form response for ${fullName} (ID: ${personId}) processed and added to 'Service Attendance' sheet.`);
  // Call to update Service Stats after a new form response is processed
  populateServiceStatsSheet();
}

// --- Service Stats Functions ---

/**
 * Creates or recreates the "Service Stats" sheet.
 */
function createServiceStatsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let statsSheet = ss.getSheetByName("Service Stats");
  if (statsSheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Sheet Already Exists',
      'Service Stats sheet already exists. Do you want to recreate it? This will clear all existing data.',
      ui.ButtonSet.YES_NO
    );
    if (response === ui.Button.YES) {
      ss.deleteSheet(statsSheet);
    } else {
      return; // User chose not to recreate
    }
  }
  statsSheet = ss.insertSheet("Service Stats");
  setupServiceStatsSheetLayout(statsSheet);
  populateServiceStatsSheet(statsSheet); // Populate immediately after creation
  Logger.log("✅ Service Stats sheet created successfully!");
  SpreadsheetApp.getUi().alert(
    'Service Stats Sheet Created!',
    'The "Service Stats" sheet has been created and populated with service attendance data.\n\n' +
    'You can refresh this data at any time from the "📋 Sunday Check-in" menu -> "Generate Service Stats Report".',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Sets up the initial layout and headers for the "Service Stats" sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to format.
 */
function setupServiceStatsSheetLayout(sheet) {
  sheet.clear();
  sheet.getRange("A1").setValue("📊 SERVICE ATTENDANCE STATISTICS").setFontSize(16).setFontWeight("bold");
  sheet.getRange("A1:K1").merge().setHorizontalAlignment("center");

  const headers = [
    "Person ID", "Full Name", "First Name", "Last Name",
    "Services This Quarter", "Services This Month", "Volunteer Count",
    "Last Attended Date", "Last Service Name", "Total Services Attended",
    "Activity Level"
  ];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#e3f2fd").setFontColor("#202124");

  // Adjust column widths for readability
  sheet.setColumnWidth(1, 100); // Person ID
  sheet.setColumnWidth(2, 200); // Full Name
  sheet.setColumnWidth(3, 120); // First Name
  sheet.setColumnWidth(4, 140); // Last Name
  sheet.setColumnWidth(5, 160); // Services This Quarter
  sheet.setColumnWidth(6, 150); // Services This Month
  sheet.setColumnWidth(7, 130); // Volunteer Count
  sheet.setColumnWidth(8, 160); // Last Attended Date
  sheet.setColumnWidth(9, 160); // Last Service Name (Column I)
  sheet.setColumnWidth(10, 160); // Total Services Attended (Column J)
  sheet.setColumnWidth(11, 120); // Activity Level (Column K)

  sheet.setFrozenRows(2); // Freeze header row
  Logger.log("✅ Service Stats sheet layout created.");
}

/**
 * Calculates attendance statistics from the "Service Attendance" sheet.
 * @returns {Array<Array<any>>} An array of arrays containing summarized attendance statistics per individual.
 */
function calculateServiceStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const serviceAttendanceSheet = ss.getSheetByName("Service Attendance");
  if (!serviceAttendanceSheet) {
    Logger.log("❌ 'Service Attendance' sheet not found for calculateServiceStats. Cannot calculate.");
    SpreadsheetApp.getUi().alert("Error", "'Service Attendance' sheet not found. Cannot generate Service Stats.", SpreadsheetApp.getUi().ButtonSet.OK);
    return [];
  }

  const serviceData = serviceAttendanceSheet.getDataRange().getValues();
  if (serviceData.length < 2) { // Only headers
    Logger.log("No data found in 'Service Attendance' sheet for statistics calculation.");
    return [];
  }

  const directoryMap = getDirectoryDataMap(); // For getting First/Last Names

  const now = new Date();
  const currentMonth = now.getMonth();
  const currentQuarter = Math.floor(currentMonth / 3);
  const currentYear = now.getFullYear();

  // Map to group attendance entries by Person ID
  const groupedById = new Map(); // Key: Person ID (string), Value: Array of attendance record objects

  // Column indices for Service Attendance data (0-based)
  const PERSON_ID_COL_SVC = 0; // Column A
  const FULL_NAME_COL_SVC = 1; // Column B
  const SERVICE_DATE_COL_SVC = 4; // Column E
  const NOTES_COL_SVC = 7; // Column H (Check if "Volunteer" or "Visitor" can be derived from here)

  for (let i = 1; i < serviceData.length; i++) { // Start from 1 to skip headers
    const row = serviceData[i];
    const personId = String(row[PERSON_ID_COL_SVC] || "").trim();
    const fullName = String(row[FULL_NAME_COL_SVC] || "").trim();
    const serviceDate = getDateValue(row[SERVICE_DATE_COL_SVC]);
    const notes = String(row[NOTES_COL_SVC] || "").toLowerCase(); // Check notes for volunteer keyword

    if (!personId || !fullName || !serviceDate) {
      Logger.log(`⚠️ Skipping row ${i + 1} in Service Attendance due to missing ID, Name, or Date.`);
      continue;
    }

    // Determine Last Service Name based on day of week of serviceDate
    let lastServiceName = "Unknown Service";
    const dayOfWeek = serviceDate.getDay(); // 0 for Sunday, 1 for Monday, ..., 5 for Friday, 6 for Saturday
    const daysOfWeek = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
    if (dayOfWeek >= 0 && dayOfWeek <= 6) {
        lastServiceName = daysOfWeek[dayOfWeek] + " Service";
    }

    // Determine if volunteer based on notes or other role column if it exists
    const isVolunteer = notes.includes("volunteer"); // Case-insensitive check

    const record = {
      personId: personId,
      fullName: fullName,
      date: serviceDate,
      month: serviceDate.getMonth(),
      quarter: Math.floor(serviceDate.getMonth() / 3),
      year: serviceDate.getFullYear(),
      lastServiceName: lastServiceName,
      isVolunteer: isVolunteer
    };

    if (!groupedById.has(personId)) {
      groupedById.set(personId, []);
    }
    groupedById.get(personId).push(record);
  }

  const summary = [];

  groupedById.forEach((records, personId) => {
    let servicesThisMonth = 0;
    let servicesThisQuarter = 0;
    let totalServicesAttended = 0; // Total count of services attended (not unique)
    let volunteerCount = 0;

    let lastAttendedDate = null;
    let lastServiceNameForSummary = ''; // The last service name for the summary row
    let personFullName = '';
    let personFirstName = '';
    let personLastName = '';
    let activityLevel = ''; // New variable for Activity Level

    // Sort records for this person by date in descending order to easily find the latest
    records.sort((a, b) => b.date.getTime() - a.date.getTime());

    // Get info from the most recent record
    if (records.length > 0) {
      const mostRecentRecord = records[0];
      personFullName = mostRecentRecord.fullName;
      lastAttendedDate = mostRecentRecord.date;
      lastServiceNameForSummary = mostRecentRecord.lastServiceName;

      // Try to get First/Last Name from Directory for more accuracy
      const dirEntry = getDirectoryDataMap().get(personFullName.toUpperCase());
      if (dirEntry) {
        personFirstName = dirEntry.firstName;
        personLastName = dirEntry.lastName;
      } else {
        // Fallback to deriving from full name if not in directory
        const nameParts = personFullName.split(/\s+/);
        personFirstName = nameParts[0] || "";
        personLastName = nameParts.length > 1 ? nameParts.slice(1).join(" ") : "";
      }
    }


    // Calculate counts for month, quarter, total services, and volunteer count for current year
    records.forEach(r => {
      totalServicesAttended++; // Each record is one service attendance

      if (r.year === currentYear) { // Only count for current year's month/quarter/volunteer
          if (r.month === currentMonth) {
            servicesThisMonth++;
          }
          if (r.quarter === currentQuarter) {
            servicesThisQuarter++;
          }
          if (r.isVolunteer) {
            volunteerCount++;
          }
      }
    });

    // Calculate Activity Level based on servicesThisQuarter
    if (servicesThisQuarter >= 12) {
      activityLevel = "Core";
    } else if (servicesThisQuarter >= 3) {
      activityLevel = "Active";
    } else {
      activityLevel = "Inactive";
    }

    // Add the calculated statistics for this individual to the summary array
    // The order here must match the columns written to in setupServiceStatsSheetLayout (11 columns A-K)
    summary.push([
      personId,
      personFullName,
      personFirstName,
      personLastName,
      servicesThisQuarter,
      servicesThisMonth,
      volunteerCount,
      lastAttendedDate,
      lastServiceNameForSummary,
      totalServicesAttended,
      activityLevel // Populated Activity Level
    ]);
  });

  Logger.log("✅ Service stats calculated for: " + summary.length + " individuals.");
  return summary;
}

/**
 * Populates the "Service Stats" sheet with calculated attendance statistics.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} targetSheet The "Service Stats" sheet.
 */
function populateServiceStatsSheet(targetSheet = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!targetSheet) {
    targetSheet = ss.getSheetByName("Service Stats");
    if (!targetSheet) {
      SpreadsheetApp.getUi().alert("Error", "'Service Stats' sheet not found. Please create it first.", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
  }

  const serviceStatsData = calculateServiceStats();

  // Clear existing data below headers
  const lastRow = targetSheet.getLastRow();
  if (lastRow > 2) { // Assuming headers are in row 1 and 2
    targetSheet.getRange(3, 1, lastRow - 2, targetSheet.getMaxColumns()).clearContent().clearFormat();
  }

  if (serviceStatsData.length > 0) {
    targetSheet.getRange(3, 1, serviceStatsData.length, serviceStatsData[0].length).setValues(serviceStatsData);
    // Format date column (Last Attended Date - Column H, index 7)
    targetSheet.getRange(3, 8, serviceStatsData.length, 1).setNumberFormat("MM/dd/yyyy");
    Logger.log(`✅ Service Stats sheet populated with ${serviceStatsData.length} entries.`);
  } else {
    Logger.log("No service statistics to populate.");
    SpreadsheetApp.getUi().alert("Info", "No service attendance data found to generate statistics.", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// --- Shared Helper Functions (Defined in both projects for full independence) ---

/**
 * RESOLVES a Person ID and associated details (First Name, Last Name, Email)
 * by checking multiple sources in a prioritized order.
 * This is the central source for Person ID lookup across all adding functions.
 *
 * @param {string} fullName The full name of the person to resolve.
 * @returns {object} An object containing { id, firstName, lastName, email }. Generates new ID if not found.
 */
function resolvePersonIdAndDetails(fullName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); // Get spreadsheet within the function
  const normalizedFullName = String(fullName || "").trim().toUpperCase();
  let personId = "";
  let firstName = "";
  let lastName = "";
  let email = "";

  Logger.log(`[resolve] Attempting to resolve ID for: ${fullName} (normalized: ${normalizedFullName})`);

  const directoryMap = getDirectoryDataMap();
  const serviceAttendanceIdMap = getLocalSheetIdMap("Service Attendance", 1, 2);
  const eventAttendanceIdMap = getLocalSheetIdMap("Event Attendance", 1, 2);
  const sundayRegMap = getLocalSheetIdMap("Sunday Registration", 1, 2);
  const eventRegMap = getLocalSheetIdMap("Event Registration", 1, 2);
  const sundayServiceFormMap = getLocalSheetIdMap("Sunday Service", 1, 2);


  const directoryEntry = directoryMap.get(normalizedFullName);
  const serviceEntryId = serviceAttendanceIdMap.get(normalizedFullName);
  const eventEntryId = eventAttendanceIdMap.get(normalizedFullName);
  const sundayRegExistingId = sundayRegMap.get(normalizedFullName);
  const eventRegExistingId = eventRegMap.get(normalizedFullName);
  const sundayServiceFormExistingId = sundayServiceFormMap.get(normalizedFullName);


  // Priority Order: Directory -> Service Attendance -> Event Attendance -> Sunday Reg -> Event Reg -> Sunday Service Form -> Generate New
  if (directoryEntry && directoryEntry.id) {
    personId = directoryEntry.id;
    firstName = directoryEntry.firstName;
    lastName = directoryEntry.lastName;
    email = directoryEntry.email;
    Logger.log(`  [resolve] -> ID found in Directory: ${personId}`);
  } else if (serviceEntryId) {
    personId = serviceEntryId;
    if (directoryEntry) { firstName = directoryEntry.firstName || firstName; lastName = directoryEntry.lastName || lastName; email = directoryEntry.email || email;}
    Logger.log(`  [resolve] -> ID found in Service Attendance: ${personId}`);
  } else if (eventEntryId) {
    personId = eventEntryId;
    if (directoryEntry) { firstName = directoryEntry.firstName || firstName; lastName = directoryEntry.lastName || lastName; email = directoryEntry.email || email;}
    Logger.log(`  [resolve] -> ID found in Event Attendance: ${personId}`);
  } else if (sundayRegExistingId) {
      personId = sundayRegExistingId;
      if (directoryEntry) { firstName = directoryEntry.firstName || firstName; lastName = directoryEntry.lastName || lastName; email = directoryEntry.email || email;}
      Logger.log(`  [resolve] -> ID found in 'Sunday Registration' list: ${personId}`);
  } else if (eventRegExistingId) {
      personId = eventRegExistingId;
      if (directoryEntry) { firstName = directoryEntry.firstName || firstName; lastName = directoryEntry.lastName || lastName; email = directoryEntry.email || email;}
      Logger.log(`  [resolve] -> ID found in 'Event Registration' list: ${personId}`);
  } else if (sundayServiceFormExistingId) {
      personId = sundayServiceFormExistingId;
      if (directoryEntry) { firstName = directoryEntry.firstName || firstName; lastName = directoryEntry.lastName || lastName; email = directoryEntry.email || email;}
      Logger.log(`  [resolve] -> ID found in 'Sunday Service' form responses: ${personId}`);
  } else {
    // If not found in any existing source, generate a new ID
    let currentHighestOverallId = Math.max(
      findHighestIdInDirectory(),
      findHighestIdInLocalSheets(LOCAL_ID_SHEETS) // Uses the global LOCAL_ID_SHEETS defined at top of this project
    );
    currentHighestOverallId++;
    personId = String(currentHighestOverallId);
    Logger.log(`  [resolve] -> Generated NEW ID: ${personId} (not found in any existing source).`);
  }

  // Fallback for first/last name if not found from Directory but full name exists
  if (!firstName && !lastName && fullName) {
    const nameParts = fullName.split(/\s+/);
    firstName = nameParts[0] || "";
    lastName = nameParts.length > 1 ? nameParts.slice(1).join(" ") : "";
    Logger.log(`  [resolve] -> Derived first/last name from full name: ${firstName} ${lastName}`);
  }

  return { id: personId, firstName: firstName, lastName: lastName, email: email };
}


/**
 * Finds the highest numerical ID in a given array of sheet names within the current spreadsheet.
 * Used to ensure new IDs are unique and incrementing.
 * @param {string[]} sheetNamesArray An array of sheet names to search for IDs.
 * @returns {number} The highest ID found, or 0 if none.
 */
function findHighestIdInLocalSheets(sheetNamesArray) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let highestId = 0;
    sheetNamesArray.forEach(sheetName => {
        const sheet = ss.getSheetByName(sheetName);
        if (sheet) {
            const lastRow = sheet.getLastRow();
            if (lastRow >= 1) {
                let startDataRow = 1;
                if ((sheetName === "Sunday Registration" || sheetName === "Event Registration") && lastRow >= 6) {
                    startDataRow = 6;
                } else if ((sheetName === "Event Attendance" || sheetName === "Service Attendance" || sheetName === "Sunday Service") && lastRow >= 2) {
                    const headerValue = String(sheet.getRange(1, 1).getDisplayValue() || "").trim().toLowerCase();
                    if (headerValue === "person id" || headerValue === "id") {
                        startDataRow = 2;
                    }
                }
                
                if (lastRow >= startDataRow) {
                    const ids = sheet.getRange(startDataRow, 1, lastRow - startDataRow + 1, 1).getValues();
                    ids.forEach(row => {
                        const id = parseInt(row[0]);
                        if (!isNaN(id) && id > highestId) {
                            highestId = id;
                        }
                    });
                }
            }
        } else {
            Logger.log(`(Shared Helper) Sheet "${sheetName}" not found for local ID generation base.`);
        }
    });
    Logger.log(`(Shared Helper) Highest current ID found across local sheets (${sheetNamesArray.join(', ')}): ${highestId}`);
    return highestId;
}

/**
 * Finds the highest numerical ID in the external "Directory" spreadsheet.
 * @returns {number} The highest ID found, or 0 if the directory is not set or accessible.
 */
function findHighestIdInDirectory() {
  let highestId = 0;
  try {
    const props = PropertiesService.getScriptProperties();
    const directoryId = props.getProperty('DIRECTORY_SPREADSHEET_ID');
    if (!directoryId) {
      Logger.log('⚠️ (Shared Helper) DIRECTORY_SPREADSHEET_ID script property not set (for findHighestIdInDirectory).');
      return 0;
    }
    const directorySS = SpreadsheetApp.openById(directoryId);
    const directorySheet = directorySS.getSheetByName("Directory");
    if (directorySheet) {
      const lastRow = directorySheet.getLastRow();
      if (lastRow >= 2) {
        const ids = directorySheet.getRange(2, 1, lastRow - 1, 1).getValues();
        ids.forEach(row => {
          const id = parseInt(row[0]);
          if (!isNaN(id) && id > highestId) {
            highestId = id;
          }
        });
      }
      Logger.log(`(Shared Helper) Highest ID found in external Directory: ${highestId}`);
    } else {
      Logger.log('⚠️ (Shared Helper) "Directory" sheet not found in the external spreadsheet (for findHighestIdInDirectory).');
    }
  } catch (error) {
    Logger.log(`❌ (Shared Helper) Error in findHighestIdInDirectory: ${error.message}`);
  }
  return highestId;
}

/**
 * Fetches data from the external "Directory" sheet and returns it as a Map for quick lookup.
 * @returns {Map<string, object>} A map where keys are normalized full names and values are objects
 * containing id, email, firstName, and lastName.
 */
function getDirectoryDataMap() {
  const directoryDataMap = new Map();
  try {
    const props = PropertiesService.getScriptProperties();
    const directoryId = props.getProperty('DIRECTORY_SPREADSHEET_ID');
    if (!directoryId) {
      Logger.log('⚠️ (Shared Helper) DIRECTORY_SPREADSHEET_ID script property not set for getDirectoryDataMap. Please set it via the Config menu.');
      return directoryDataMap;
    }
    const directorySS = SpreadsheetApp.openById(directoryId);
    const directorySheet = directorySS.getSheetByName("Directory");
    if (directorySheet) {
      const directoryValues = directorySheet.getDataRange().getValues();
      if (directoryValues.length > 1) {
        const headers = directoryValues[0].map(h => String(h || "").trim().toLowerCase());
        const idColIndex = 0; // Column A (0-indexed) for Person ID
        const nameColIndex = 1; // Column B (0-indexed) for Full Name

        let firstNameColIndex = headers.indexOf("first name");
        if (firstNameColIndex === -1) firstNameColIndex = headers.indexOf("firstname");
        if (firstNameColIndex === -1) firstNameColIndex = 2; // Fallback to C (0-indexed 2)

        let lastNameColIndex = headers.indexOf("last name");
        if (lastNameColIndex === -1) lastNameColIndex = headers.indexOf("lastname");
        if (lastNameColIndex === -1) lastNameColIndex = 3; // Fallback to D (0-indexed 3)

        let emailColIndex = headers.indexOf("email");
        if (emailColIndex === -1) emailColIndex = 7; // Fallback to H (0-indexed 7)

        for (let i = 1; i < directoryValues.length; i++) {
          const row = directoryValues[i];
          const personId = String(row[idColIndex] || "").trim();
          const fullName = String(row[nameColIndex] || "").trim();
          if (personId && fullName) {
            const normalizedFullName = fullName.toUpperCase();
            directoryDataMap.set(normalizedFullName, {
              id: personId,
              email: emailColIndex !== -1 && row[emailColIndex] !== undefined ? String(row[emailColIndex]).trim() : "",
              firstName: firstNameColIndex !== -1 && row[firstNameColIndex] !== undefined ? String(row[firstNameColIndex]).trim() : "",
              lastName: lastNameColIndex !== -1 && row[lastNameColIndex] !== undefined ? String(row[lastNameColIndex]).trim() : "",
              originalFullName: fullName
            });
          }
        }
      }
      Logger.log(`(Shared Helper) Directory data map created with ${directoryDataMap.size} entries.`);
    } else {
      Logger.log('⚠️ (Shared Helper) "Directory" sheet not found in the external spreadsheet for getDirectoryDataMap. Please ensure the sheet name is "Directory".');
    }
  } catch (error) {
    Logger.log(`❌ (Shared Helper) Error in getDirectoryDataMap: ${error.message}. Ensure the ID is correct and you have access.`);
  }
  return directoryDataMap;
}

/**
 * Gets a map of full names to Person IDs from a specified local sheet.
 * @param {string} sheetName The name of the local sheet to check.
 * @param {number} idColNum The 1-indexed column number for Person ID.
 * @param {number} nameColNum The 1-indexed column number for Full Name.
 * @returns {Map<string, string>} A map where keys are normalized full names and values are Person IDs.
 */
function getLocalSheetIdMap(sheetName, idColNum = 1, nameColNum = 2) {
    const localIdMap = new Map();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
        const lastRow = sheet.getLastRow();
        if (lastRow >= 1) {
            const data = sheet.getRange(1, 1, lastRow, Math.max(idColNum, nameColNum)).getValues();
            let dataStartRowIndex = 0;

            if (lastRow > 1) {
                if ((sheetName === "Sunday Registration" || sheetName === "Event Registration") && lastRow >= 6) {
                    dataStartRowIndex = 5;
                } else if ((sheetName === "Event Attendance" || sheetName === "Service Attendance" || sheetName === "Sunday Service") && lastRow >= 2) {
                    const headerIdCell = String(data[0][idColNum - 1] || "").trim().toLowerCase();
                    if (headerIdCell === "person id" || headerIdCell === "id") {
                        dataStartRowIndex = 1;
                    }
                }
            }

            for (let i = dataStartRowIndex; i < data.length; i++) {
                const row = data[i];
                const personId = String(row[idColNum - 1] || "").trim();
                const fullName = String(row[nameColNum - 1] || "").trim();
                if (personId && fullName) {
                    localIdMap.set(fullName.toUpperCase(), personId);
                }
            }
        }
        Logger.log(`(Shared Helper) Local ID map created for "${sheetName}" with ${localIdMap.size} entries.`);
    } else {
        Logger.log(`⚠️ (Shared Helper) Local sheet "${sheetName}" not found for ID lookup.`);
    }
    return localIdMap;
}

/**
 * Applies alternating row formatting (zebra stripping) to the data rows.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to format.
 * @param {number} startDataRow The row number where data starts (e.g., 6).
 * @param {number} numRowsInput The number of data rows to format. If -1, it calculates based on last row.
 */
function refreshRowFormatting(sheet, startDataRow = 6, numRowsInput = -1) {
  if (!sheet) { // Try to get the active sheet if not provided, for robustness
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const sheetName = activeSheet.getName();
    if (sheetName === "Sunday Registration" || sheetName === "Event Registration") {
      sheet = activeSheet;
    } else {
      Logger.log("(Shared Helper) refreshRowFormatting called without a sheet and active sheet is not a registration sheet.");
      return;
    }
  }

  let numRowsToFormat = numRowsInput;
  if (numRowsToFormat === -1) {
    const lastSheetRowWithContent = findLastRowWithData(sheet);
    if (lastSheetRowWithContent < startDataRow) {
      Logger.log("(Shared Helper) No data rows to format in refreshRowFormatting.");
      return;
    }
    numRowsToFormat = lastSheetRowWithContent - startDataRow + 1;
  }

  if (numRowsToFormat <= 0) {
    Logger.log("(Shared Helper) Calculated numRowsToFormat is <= 0 in refreshRowFormatting.");
    return;
  }

  sheet.getRange(startDataRow, 1, numRowsToFormat, 5).clearFormat();

  for (let i = 0; i < numRowsToFormat; i++) {
    const currentRowInSheet = startDataRow + i;
    const rowRange = sheet.getRange(currentRowInSheet, 1, 1, 5);
    if (i % 2 === 1) {
      rowRange.setBackground("#f5f5f5");
    } else {
      rowRange.setBackground("white");
    }
  }
  Logger.log(`(Shared Helper) Row formatting refreshed for ${numRowsToFormat} rows starting at ${startDataRow} on sheet ${sheet.getName()}.`);
}


/**
 * Finds the last row in a sheet that contains any data. More robust than getLastRow().
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to check.
 * @returns {number} The row number of the last row with data, or 0 if empty.
 */
function findLastRowWithData(sheet) {
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) return 0;
  const maxCol = sheet.getMaxColumns();
  if (maxCol === 0) return 0;

  for (let r = lastRow; r >= 1; r--) {
    const range = sheet.getRange(r, 1, 1, maxCol);
    if (!range.isBlank()) {
      return r;
    }
  }
  return 0;
}

/**
 * Extracts the spreadsheet ID from a Google Sheet URL.
 * @param {string} url The full Google Sheet URL.
 * @returns {string} The extracted spreadsheet ID, or null if not found.
 */
function extractSpreadsheetIdFromUrl(url) {
  const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

/**
 * Helper function to safely convert a value to a Date object.
 * Returns null if conversion is not successful.
 */
function getDateValue(value) {
  if (value instanceof Date) {
    return value;
  }
  if (typeof value === 'string') {
    try {
      const date = new Date(value);
      if (!isNaN(date.getTime()) && date.getFullYear() > 1900) {
        return date;
      }
    } catch (e) { /* Fall through if parsing fails */ }
  }
  // For numbers that represent dates (e.g., from CSV import without formatting)
  if (typeof value === 'number' && value > 0 && value < 2958466) {
      try {
          const date = new Date((value - 25569) * 86400 * 1000); // Convert Excel/Sheets date serial to milliseconds
          if (!isNaN(date.getTime())) return date;
      } catch(e) { /* Fall through if error converting number to date */ }
  }
  return null;
}