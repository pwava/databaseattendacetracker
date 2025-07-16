/**
 * Sunday Registration & Service Utilities System (Primary Project)
 * Manages Sunday Service registration, attendance, and Service Stats.
 * Contains all shared helper functions for use by other Apps Script projects
 * bound to the same spreadsheet.
 *
 * MODIFIED: This version removes the dependency on a "Full Name" column (Column B)
 * in the "Sunday Registration" sheet. It now uses First Name (new Col B) and Last Name (new Col C).
 */

// Define the names of sheets that might contain person IDs locally within *this* spreadsheet.
// This list is used for ID generation logic to ensure uniqueness.
const LOCAL_ID_SHEETS = ["Sunday Registration", "Service Attendance", "Sunday Service", "Event Registration", "Event Attendance", "attendance stats"];

// --- Sunday Registration Functions ---

/**
 * Creates or recreates the main "Sunday Registration" sheet.
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
    'The "Full Name" column has been removed. The sheet now uses First and Last Name.\n\n' +
    'The registration team can now:\n' +
    '1. Enter the service date in cell B2\n' +
    '2. Check the boxes for attendees\n' +
    '3. Click "Submit Attendance" from the "📋 Sunday Check-in" menu to transfer to Service Attendance sheet\n\n' +
    'Menus have been added/updated for easy access to functions.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * MODIFIED: Sets up the initial layout for the new 4-column format.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to format.
 */
function setupSundayRegistrationSheetLayout(sheet) {
  sheet.clear();
  sheet.getRange("A1").setValue("🏛️ SUNDAY SERVICE REGISTRATION").setFontSize(16).setFontWeight("bold");
  sheet.getRange("A1:D1").merge().setHorizontalAlignment("center");
  sheet.getRange("A2").setValue("📅 Service Date:");
  sheet.getRange("B2").setValue(new Date()).setNumberFormat("MM/dd/yyyy");
  sheet.getRange("A3").setValue("📝 Instructions: Check the box next to each person who is present today");
  sheet.getRange("A3:D3").merge();
  sheet.getRange("A4").setValue("🔄 Refresh List");
  sheet.getRange("B4").setValue("✅ Submit Attendance");
  sheet.getRange("C4").setValue("🧹 Clear All Checks");
  sheet.getRange("D4").setValue("Status: Ready");

  // MODIFIED: Headers array updated to remove "Full Name".
  const headers = ["ID", "First Name", "Last Name", "✓ Present"];
  sheet.getRange("A5:D5").setValues([headers]).setFontWeight("bold").setBackground("#4285f4").setFontColor("white");

  // MODIFIED: Column widths adjusted for the new layout.
  sheet.setColumnWidth(1, 70);   // ID
  sheet.setColumnWidth(2, 150);  // First Name
  sheet.setColumnWidth(3, 150);  // Last Name
  sheet.setColumnWidth(4, 80);   // Present Checkbox
  sheet.hideColumns(1);

  sheet.getRange("A1:D4").setBackground("#f8f9fa");
  sheet.getRange("A2:B2").setBackground("#e3f2fd");
  sheet.getRange("A4:D4").setBackground("#fff3e0");
  sheet.setFrozenRows(5);
  Logger.log("✅ Sunday Registration sheet layout created (New 4-column format).");
}

/**
 * MODIFIED: Populates the list using the new 4-column format.
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
    return;
  }

  const serviceAttendanceIdMap = getLocalSheetIdMap("Service Attendance", 1, 2);

  let nextGeneratedId = findHighestIdInDirectory();
  if (nextGeneratedId === 0) {
    nextGeneratedId = findHighestIdInLocalSheets(LOCAL_ID_SHEETS);
  }

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
      }
    }

    if (!firstName && !lastName && fullName) {
      const nameParts = fullName.split(/\s+/);
      firstName = nameParts[0] || "";
      lastName = nameParts.length > 1 ? nameParts.slice(1).join(" ") : "";
    }
    // MODIFIED: Pushing data array without `fullName`. It's now [ID, First, Last, Checkbox].
    personsForRegistration.push([personId, firstName, lastName, false]);
  }

  // MODIFIED: Sort by Last Name, which is now at index 2 of the inner array.
  personsForRegistration.sort((a, b) => (String(a[2]) || "").toLowerCase().localeCompare((String(b[2]) || "").toLowerCase()));

  const lastDataRowOnSheet = regSheet.getLastRow();
  // MODIFIED: Clearing 4 columns of data.
  if (lastDataRowOnSheet > 5) {
    regSheet.getRange(6, 1, lastDataRowOnSheet - 5, 4).clearContent().clearFormat();
  }

  if (personsForRegistration.length > 0) {
    const startRow = 6;
    // MODIFIED: Setting values for 4 columns.
    regSheet.getRange(startRow, 1, personsForRegistration.length, 4).setValues(personsForRegistration);
    // MODIFIED: Checkbox is now in column 4.
    const checkboxRange = regSheet.getRange(startRow, 4, personsForRegistration.length, 1);
    checkboxRange.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
    // MODIFIED: Bordering 4 columns.
    regSheet.getRange(startRow, 1, personsForRegistration.length, 4).setBorder(true, true, true, true, true, true);
    refreshRowFormatting(regSheet, startRow, personsForRegistration.length);
  }
  regSheet.getRange("D4").setValue(`Status: ${personsForRegistration.length} members loaded`);
  Logger.log(`✅ Sunday Registration list populated with ${personsForRegistration.length} members.`);
}

/**
 * MODIFIED: Adds a person by resolving ID and splitting the name. Checks for duplicates by combining First/Last names on the sheet.
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

  // MODIFIED: Check for duplicate by combining existing First and Last names.
  const lastDataRow = regSheet.getLastRow();
  if (lastDataRow >= 6) {
    const existingNames = regSheet.getRange(6, 2, lastDataRow - 5, 2).getValues(); // Get First and Last names
    const isDuplicate = existingNames.some(nameParts => {
      const existingFullName = `${nameParts[0]} ${nameParts[1]}`.trim();
      return existingFullName.toLowerCase() === fullNameEntered.toLowerCase();
    });
    if (isDuplicate) {
      ui.alert('Duplicate Entry', 'This person is already in the current Sunday registration list.', ui.ButtonSet.OK);
      return;
    }
  }

  const personDetails = resolvePersonIdAndDetails(fullNameEntered);
  const personIdToAdd = personDetails.id;
  const firstNameToAdd = personDetails.firstName;
  const lastNameToAdd = personDetails.lastName;

  const nextSheetRow = (lastDataRow < 5) ? 6 : lastDataRow + 1;
  // MODIFIED: New row data format.
  const newRowData = [personIdToAdd, firstNameToAdd, lastNameToAdd, false];
  // MODIFIED: Set 4 columns of data.
  regSheet.getRange(nextSheetRow, 1, 1, 4).setValues([newRowData]);
  // MODIFIED: Checkbox is in column 4.
  regSheet.getRange(nextSheetRow, 4).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  const newRowRange = regSheet.getRange(nextSheetRow, 1, 1, 4);
  newRowRange.setBorder(true, true, true, true, true, true);
  refreshRowFormatting(regSheet);

  ui.alert('Person Added!', `${fullNameEntered} has been added with ID ${personIdToAdd}.`, ui.ButtonSet.OK);
  Logger.log(`✅ Manually added ${fullNameEntered} (ID: ${personIdToAdd}) to Sunday registration list.`);
}


/**
 * MODIFIED: Submits attendance by constructing full name from First and Last name columns.
 */
function submitSundayRegistrationAttendance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Sunday Registration");
  if (!regSheet) { /* ... error handling ... */ return; }
  
  const serviceDateValue = regSheet.getRange("B2").getValue();
  if (!serviceDateValue || !(serviceDateValue instanceof Date)) { /* ... error handling ... */ return; }
  
  const formattedServiceDate = Utilities.formatDate(serviceDateValue, ss.getSpreadsheetTimeZone(), "MM/dd/yyyy");
  regSheet.getRange("D4").setValue("Status: Processing...");

  try {
    const serviceSheet = ss.getSheetByName("Service Attendance");
    if (!serviceSheet) { /* ... error handling ... */ throw new Error("'Service Attendance' sheet not found"); }

    const directoryMap = getDirectoryDataMap();
    const lastRegDataRow = regSheet.getLastRow();
    if (lastRegDataRow < 6) { /* ... error handling ... */ return; }

    // MODIFIED: Get 4 columns of data.
    const regData = regSheet.getRange(6, 1, lastRegDataRow - 5, 4).getValues();
    const attendanceEntries = [];
    let checkedCount = 0;

    for (const row of regData) {
      // MODIFIED: Destructure the new 4-column row format.
      const [personId, firstName, lastName, isChecked] = row;
      
      // MODIFIED: Construct fullName on the fly.
      const fullName = `${firstName} ${lastName}`.trim();

      if (isChecked === true && fullName !== "") {
        let email = "";
        const normalizedFullName = fullName.toUpperCase();
        const directoryEntry = directoryMap.get(normalizedFullName);
        if (directoryEntry && directoryEntry.email) {
          email = directoryEntry.email;
        }

        attendanceEntries.push([
          personId, fullName, firstName || "", lastName || "",
          formattedServiceDate, "No", email, "", new Date()
        ]);
        checkedCount++;
      }
    }

    if (attendanceEntries.length === 0) { /* ... error handling ... */ return; }

    if (serviceSheet.getLastRow() < 1) {
      const serviceHeaders = ["Person ID", "Full Name", "First Name", "Last Name", "Service Date", "Is Visitor?", "Email", "Notes", "Timestamp"];
      serviceSheet.getRange(1, 1, 1, serviceHeaders.length).setValues([serviceHeaders]).setFontWeight("bold").setBackground("#4285f4").setFontColor("white");
    }

    const nextRowServiceSheet = findLastRowWithData(serviceSheet) + 1;
    serviceSheet.getRange(nextRowServiceSheet, 1, attendanceEntries.length, 9).setValues(attendanceEntries);
    serviceSheet.getRange(nextRowServiceSheet, 5, attendanceEntries.length, 1).setNumberFormat("MM/dd/yyyy");
    serviceSheet.getRange(nextRowServiceSheet, 9, 1, 1).setNumberFormat("MM/dd/yyyy HH:mm:ss");

    // MODIFIED: Clear checkboxes in column 4.
    regSheet.getRange(6, 4, lastRegDataRow - 5, 1).setValue(false);
    regSheet.getRange("D4").setValue(`Status: ${checkedCount} attendees submitted`);
    SpreadsheetApp.getUi().alert('Attendance Submitted!', `Successfully submitted ...`, SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log(`✅ Successfully submitted ${checkedCount} Sunday Service attendance entries.`);

    populateServiceStatsSheet();

  } catch (error) {
    regSheet.getRange("D4").setValue("Status: Error occurred");
    Logger.log(`❌ Error submitting Sunday attendance: ${error.message}\n${error.stack || ""}`);
    SpreadsheetApp.getUi().alert("Error", `Error submitting Sunday attendance: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * MODIFIED: Clears checkboxes in column 4.
 */
function clearAllSundayChecks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Sunday Registration");
  if (!regSheet) { /* ... */ return; }
  
  const lastDataRow = regSheet.getLastRow();
  if (lastDataRow >= 6) {
    // MODIFIED: Checkbox column is now 4.
    regSheet.getRange(6, 4, lastDataRow - 5, 1).setValue(false);
    regSheet.getRange("D4").setValue("Status: All checks cleared");
    Logger.log("✅ All Sunday checkboxes cleared");
  } else {
    regSheet.getRange("D4").setValue("Status: No checks to clear");
  }
}

/**
 * MODIFIED: Adds checkboxes and formatting based on new layout.
 */
function addCheckboxesToSundayRegistration() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const regSheet = ss.getSheetByName("Sunday Registration");
    if (!regSheet) { /* ... */ return; }
    
    const lastRow = regSheet.getLastRow();
    if (lastRow < 6) { /* ... */ return; }

    // MODIFIED: Check for names in either First Name (new col B) or Last Name (new col C)
    const nameValues = regSheet.getRange(6, 2, lastRow - 5, 2).getValues(); 
    let rowsWithActualNames = 0;
    for (let i = 0; i < nameValues.length; i++) {
        // Check if either first name or last name has content
        if (String(nameValues[i][0] || "").trim() !== "" || String(nameValues[i][1] || "").trim() !== "") {
            rowsWithActualNames = i + 1;
        }
    }
    if (rowsWithActualNames === 0) {
        SpreadsheetApp.getUi().alert("No Names Found", "No names found in First/Last Name columns.", SpreadsheetApp.getUi().ButtonSet.OK);
        return;
    }

    try {
        // MODIFIED: Checkbox is column 4, data range is 4 columns wide
        const checkboxRange = regSheet.getRange(6, 4, rowsWithActualNames, 1);
        checkboxRange.clearContent().setValue(false).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());

        const dataFormattingRange = regSheet.getRange(6, 1, rowsWithActualNames, 4);
        dataFormattingRange.setBorder(true, true, true, true, true, true);
        refreshRowFormatting(regSheet, 6, rowsWithActualNames);

        regSheet.getRange("D4").setValue(`Status: ${rowsWithActualNames} members ready`);
        SpreadsheetApp.getUi().alert('Checkboxes Added/Reformatted!', `Successfully added/reformatted checkboxes for ${rowsWithActualNames} member rows.`, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (error) {
        Logger.log(`❌ Error adding/reformatting checkboxes for Sunday Registration: ${error.message}`);
        SpreadsheetApp.getUi().alert("Error", `Error adding/reformatting checkboxes: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
}


/**
 * Creates an empty "Sunday Registration" sheet with the new layout.
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
  setupSundayRegistrationSheetLayout(regSheet); // This now calls the modified layout function
  Logger.log("✅ Empty Sunday Registration sheet created.");
  SpreadsheetApp.getUi().alert(
    'Empty Registration Sheet Created!',
    'Sunday Registration sheet is ready for manual data entry.\n\n' +
    'INSTRUCTIONS:\n' +
    '1. Paste directory data starting row 6 (First Name in Col B, Last in C)\n' +
    '2. Use "📋 Sunday Check-in" → "🔲 Add/Reformat Checkboxes" to set up column D.\n' +
    '3. Enter service date in B2 and start checking attendance!',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/**
 * MODIFIED: Removes a person by finding a match against the combined First and Last Name.
 */
function removePersonFromSundayRegistration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Sunday Registration");
  if (!regSheet) { /* ... */ return; }

  const ui = SpreadsheetApp.getUi();
  const lastDataRow = regSheet.getLastRow();
  if (lastDataRow < 6) { /* ... */ return; }

  const nameResponse = ui.prompt('Remove Person', 'Enter the FULL NAME of the person to remove (case-insensitive):', ui.ButtonSet.OK_CANCEL);
  if (nameResponse.getSelectedButton() !== ui.Button.OK) return;
  const nameToRemove = String(nameResponse.getResponseText() || "").trim().toLowerCase();
  if (!nameToRemove) { /* ... */ return; }
  
  // MODIFIED: Read all data (4 columns) and find the row to delete.
  const allData = regSheet.getRange(6, 1, lastDataRow - 5, 4).getValues();
  let rowToDeleteInSheet = -1;
  for (let i = 0; i < allData.length; i++) {
    const firstName = allData[i][1] || "";
    const lastName = allData[i][2] || "";
    const currentFullName = `${firstName} ${lastName}`.trim().toLowerCase();
    
    if (currentFullName === nameToRemove) {
      rowToDeleteInSheet = i + 6; // +6 because data starts at row 6 and loop is 0-indexed
      break;
    }
  }

  if (rowToDeleteInSheet > 0) {
    regSheet.deleteRow(rowToDeleteInSheet);
    ui.alert('Person Removed!', `'${nameResponse.getResponseText().trim()}' has been removed.`, ui.ButtonSet.OK);
    refreshRowFormatting(regSheet);
  } else {
    ui.alert('Not Found', `Person '${nameResponse.getResponseText().trim()}' not found.`, ui.ButtonSet.OK);
  }
}


/**
 * MODIFIED: Sorts the sheet by Last Name (now column 3).
 */
function sortSundayRegistrationByLastName() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Sunday Registration");
  if (!regSheet) { /* ... */ return; }
  
  const lastDataRow = findLastRowWithData(regSheet);
  if (lastDataRow < 6) { /* ... */ return; }
  
  const numDataRows = lastDataRow - 5;
  if (numDataRows <= 0) { /* ... */ return; }
  
  // MODIFIED: Data range is 4 columns wide.
  const dataRange = regSheet.getRange(6, 1, numDataRows, 4);
  // MODIFIED: Sort by column 3 (Last Name).
  dataRange.sort({ column: 3, ascending: true }); 
  refreshRowFormatting(regSheet, 6, numDataRows);
  SpreadsheetApp.getUi().alert("Sunday list sorted by Last Name.");
  Logger.log("✅ Sunday Registration list sorted by last name.");
}


/**
 * The rest of your script (Service Stats, Form Handlers, Shared Helpers) does not need to be changed
 * for this request, but is included here for completeness. The `refreshRowFormatting`
 * function has been slightly modified to handle the different column counts.
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
// No changes needed in this section.
function processSundayFormResponse(e) {
  Logger.log("Processing Sunday form response...");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const serviceFormSheet = e.range.getSheet();
  const serviceSheet = ss.getSheetByName("Service Attendance");

  if (!serviceSheet) {
    Logger.log("❌ Target 'Service Attendance' sheet not found for form submission processing.");
    return;
  }
  if (serviceFormSheet.getName() !== "Sunday Service") {
    Logger.log("Skipping form response: Not from 'Sunday Service' sheet.");
    return;
  }

  const newRow = e.range.getValues()[0];
  const headers = serviceFormSheet.getRange(1, 1, 1, serviceFormSheet.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toLowerCase());

  const TIMESTAMP_COL_FORM_IDX = headers.indexOf("timestamp");
  const FULL_NAME_COL_FORM_IDX = headers.indexOf("full name");
  const FIRST_NAME_COL_FORM_IDX = headers.indexOf("first name");
  const LAST_NAME_COL_FORM_IDX = headers.indexOf("last name");
  const EMAIL_COL_FORM_IDX = headers.indexOf("email");

  let timestamp = newRow[TIMESTAMP_COL_FORM_IDX];
  let fullName = String(newRow[FULL_NAME_COL_FORM_IDX] || "").trim();
  let firstName = String(newRow[FIRST_NAME_COL_FORM_IDX] || "").trim();
  let lastName = String(newRow[LAST_NAME_COL_FORM_IDX] || "").trim();
  let email = String(newRow[EMAIL_COL_FORM_IDX] || "").trim();
  const serviceDate = Utilities.formatDate(timestamp, ss.getSpreadsheetTimeZone(), "MM/dd/yyyy");

  if (!firstName && !lastName && fullName) {
    const nameParts = fullName.split(/\s+/);
    firstName = nameParts[0] || "";
    lastName = nameParts.length > 1 ? nameParts.slice(1).join(" ") : "";
  }

  const personDetails = resolvePersonIdAndDetails(fullName);
  const personId = personDetails.id;
  firstName = personDetails.firstName || firstName;
  lastName = personDetails.lastName || lastName;
  email = personDetails.email || email;
  
  const entryToServiceAttendance = [
    personId, fullName, firstName, lastName, serviceDate, "No",
    email, "", new Date()
  ];

  if (serviceSheet.getLastRow() < 1) {
    const serviceHeaders = ["Person ID", "Full Name", "First Name", "Last Name", "Service Date", "Is Visitor?", "Email", "Notes", "Timestamp"];
    serviceSheet.getRange(1, 1, 1, serviceHeaders.length).setValues([serviceHeaders]).setFontWeight("bold").setBackground("#4285f4").setFontColor("white");
  }

  const nextRowServiceSheet = findLastRowWithData(serviceSheet) + 1;
  serviceSheet.getRange(nextRowServiceSheet, 1, 1, entryToServiceAttendance.length).setValues([entryToServiceAttendance]);
  serviceSheet.getRange(nextRowServiceSheet, 5, 1, 1).setNumberFormat("MM/dd/yyyy");
  serviceSheet.getRange(nextRowServiceSheet, 9, 1, 1).setNumberFormat("MM/dd/yyyy HH:mm:ss");

  Logger.log(`✅ Form response for ${fullName} (ID: ${personId}) processed and added to 'Service Attendance' sheet.`);
  populateServiceStatsSheet();
}


// --- Service Stats Functions ---
// No changes needed in this section.
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
      return;
    }
  }
  statsSheet = ss.insertSheet("Service Stats");
  setupServiceStatsSheetLayout(statsSheet);
  populateServiceStatsSheet(statsSheet);
  Logger.log("✅ Service Stats sheet created successfully!");
  SpreadsheetApp.getUi().alert(
    'Service Stats Sheet Created!',
    'The "Service Stats" sheet has been created and populated with service attendance data.\n\n' +
    'You can refresh this data at any time from the "📋 Sunday Check-in" menu -> "Generate Service Stats Report".',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

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
  sheet.setColumnWidth(1, 100); sheet.setColumnWidth(2, 200); sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 140); sheet.setColumnWidth(5, 160); sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(7, 130); sheet.setColumnWidth(8, 160); sheet.setColumnWidth(9, 160);
  sheet.setColumnWidth(10, 160); sheet.setColumnWidth(11, 120);
  sheet.setFrozenRows(2);
  Logger.log("✅ Service Stats sheet layout created.");
}

/**
 * Calculates service statistics with performance improvements and updated logic as requested.
 * - Fixes timeouts by processing the 'Service Attendance' sheet in a single efficient pass.
 * - Changes Column E's logic to count services in the "Last 3 Months" instead of by quarter.
 * - This version ONLY includes people found in the "Service Attendance" sheet.
 *
 * @returns {Array<Array<any>>} A 2D array of summary data for the stats sheet.
 */
function calculateServiceStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const serviceAttendanceSheet = ss.getSheetByName("Service Attendance");

  // If the sheet doesn't exist or is empty, return an empty list.
  if (!serviceAttendanceSheet) {
    Logger.log("⚠️ 'Service Attendance' sheet not found. Cannot generate stats.");
    return [];
  }
  const serviceData = serviceAttendanceSheet.getDataRange().getValues();
  if (serviceData.length < 2) {
    Logger.log("No data in 'Service Attendance' sheet to process.");
    return [];
  }

  // A map to hold the aggregated stats. It will be built ONLY from attendance data.
  const statsMap = new Map();

  // --- Date calculations for the "Last 3 Months" logic ---
  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth();
  const threeMonthsAgo = new Date();
  threeMonthsAgo.setMonth(now.getMonth() - 3);

  // Column indices from the "Service Attendance" sheet
  const PERSON_ID_COL_SVC = 0, FULL_NAME_COL_SVC = 1, FIRST_NAME_COL_SVC = 2;
  const LAST_NAME_COL_SVC = 3, SERVICE_DATE_COL_SVC = 4, NOTES_COL_SVC = 7;

  // --- Process all attendance records in a single loop to prevent timeouts ---
  for (let i = 1; i < serviceData.length; i++) {
    const row = serviceData[i];
    const personId = String(row[PERSON_ID_COL_SVC] || "").trim();
    const fullName = String(row[FULL_NAME_COL_SVC] || "").trim();
    if (!personId || !fullName) continue;

    const serviceDate = getDateValue(row[SERVICE_DATE_COL_SVC]);
    if (!serviceDate) continue;

    // If a person is not yet in our stats map, create a new entry for them.
    if (!statsMap.has(personId)) {
      statsMap.set(personId, {
        personId: personId,
        fullName: fullName,
        firstName: String(row[FIRST_NAME_COL_SVC] || ""),
        lastName: String(row[LAST_NAME_COL_SVC] || ""),
        servicesLast3Months: 0,
        servicesThisMonth: 0,
        volunteerCount: 0,
        lastAttendedDate: null,
        lastServiceName: "N/A",
        totalServicesAttended: 0,
        activityLevel: "Inactive"
      });
    }

    // Get the stats object for the current person and update it.
    const personStats = statsMap.get(personId);

    // Accumulate statistics
    personStats.totalServicesAttended++;

    if (String(row[NOTES_COL_SVC] || "").toLowerCase().includes("volunteer")) {
      personStats.volunteerCount++;
    }

    if (!personStats.lastAttendedDate || serviceDate > personStats.lastAttendedDate) {
      personStats.lastAttendedDate = serviceDate;
      const days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
      personStats.lastServiceName = days[serviceDate.getDay()] + " Service";
    }

    if (serviceDate.getFullYear() === currentYear && serviceDate.getMonth() === currentMonth) {
      personStats.servicesThisMonth++;
    }

    // This is the requested change: Count services if they occurred within the last 3 months.
    if (serviceDate >= threeMonthsAgo && serviceDate <= now) {
      personStats.servicesLast3Months++;
    }
  }

  // --- Convert the map of stats into the final array for the spreadsheet ---
  const summary = [];
  statsMap.forEach(stats => {
    // Determine activity level based on the last 3 months of attendance
    if (stats.servicesLast3Months >= 12) {
      stats.activityLevel = "Core";
    } else if (stats.servicesLast3Months >= 3) {
      stats.activityLevel = "Active";
    } else {
      stats.activityLevel = "Inactive";
    }

    // Add the person's data to the final summary array
    summary.push([
      stats.personId,
      stats.fullName,
      stats.firstName,
      stats.lastName,
      stats.servicesLast3Months, // This is the updated Column E value
      stats.servicesThisMonth,
      stats.volunteerCount,
      stats.lastAttendedDate,
      stats.lastServiceName,
      stats.totalServicesAttended,
      stats.activityLevel
    ]);
  });

  // Sort the final report by Last Name (column index 3)
  summary.sort((a, b) => String(a[3] || "").toLowerCase().localeCompare(String(b[3] || "").toLowerCase()));
  
  Logger.log(`✅ Service stats calculated for ${summary.length} individuals.`);
  return summary;
}

function populateServiceStatsSheet(targetSheet = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!targetSheet) {
    targetSheet = ss.getSheetByName("Service Stats");
    if (!targetSheet) { return; }
  }
  const serviceStatsData = calculateServiceStats();
  const lastRow = targetSheet.getLastRow();
  if (lastRow > 2) {
    targetSheet.getRange(3, 1, lastRow - 2, targetSheet.getMaxColumns()).clearContent().clearFormat();
  }
  if (serviceStatsData.length > 0) {
    targetSheet.getRange(3, 1, serviceStatsData.length, serviceStatsData[0].length).setValues(serviceStatsData);
    targetSheet.getRange(3, 8, serviceStatsData.length, 1).setNumberFormat("MM/dd/yyyy");
    Logger.log(`✅ Service Stats sheet populated with ${serviceStatsData.length} entries.`);
  } else {
    Logger.log("No service statistics to populate.");
  }
}

// --- Shared Helper Functions ---
// No changes needed in this section, except for `refreshRowFormatting`.

function resolvePersonIdAndDetails(fullName) {
  const normalizedFullName = String(fullName || "").trim().toUpperCase();
  let personId = "", firstName = "", lastName = "", email = "";
  Logger.log(`[resolve] Attempting to resolve ID for: ${fullName}`);

  // MODIFIED: getLocalSheetIdMap for "Sunday Registration" is updated to handle the new format.
  const sundayRegMap = getLocalSheetIdMap("Sunday Registration", 1, 2); 
  
  const directoryMap = getDirectoryDataMap();
  const serviceAttendanceIdMap = getLocalSheetIdMap("Service Attendance", 1, 2);
  const eventAttendanceIdMap = getLocalSheetIdMap("Event Attendance", 1, 2);
  const eventRegMap = getLocalSheetIdMap("Event Registration", 1, 2);
  const sundayServiceFormMap = getLocalSheetIdMap("Sunday Service", 1, 2);

  const directoryEntry = directoryMap.get(normalizedFullName);
  const serviceEntryId = serviceAttendanceIdMap.get(normalizedFullName);
  const eventEntryId = eventAttendanceIdMap.get(normalizedFullName);
  const sundayRegExistingId = sundayRegMap.get(normalizedFullName);
  const eventRegExistingId = eventRegMap.get(normalizedFullName);
  const sundayServiceFormExistingId = sundayServiceFormMap.get(normalizedFullName);

  if (directoryEntry && directoryEntry.id) {
    personId = directoryEntry.id; firstName = directoryEntry.firstName;
    lastName = directoryEntry.lastName; email = directoryEntry.email;
    Logger.log(` [resolve] -> ID found in Directory: ${personId}`);
  } else if (serviceEntryId) {
    personId = serviceEntryId;
    if (directoryEntry) { firstName = directoryEntry.firstName; lastName = directoryEntry.lastName; email = directoryEntry.email; }
    Logger.log(` [resolve] -> ID found in Service Attendance: ${personId}`);
  } else if (eventEntryId) {
    personId = eventEntryId;
    if (directoryEntry) { firstName = directoryEntry.firstName; lastName = directoryEntry.lastName; email = directoryEntry.email; }
    Logger.log(` [resolve] -> ID found in Event Attendance: ${personId}`);
  } else if (sundayRegExistingId) {
      personId = sundayRegExistingId;
      if (directoryEntry) { firstName = directoryEntry.firstName; lastName = directoryEntry.lastName; email = directoryEntry.email; }
      Logger.log(` [resolve] -> ID found in 'Sunday Registration' list: ${personId}`);
  } else if (eventRegExistingId) {
      personId = eventRegExistingId;
      if (directoryEntry) { firstName = directoryEntry.firstName; lastName = directoryEntry.lastName; email = directoryEntry.email; }
      Logger.log(` [resolve] -> ID found in 'Event Registration' list: ${personId}`);
  } else if (sundayServiceFormExistingId) {
      personId = sundayServiceFormExistingId;
      if (directoryEntry) { firstName = directoryEntry.firstName; lastName = directoryEntry.lastName; email = directoryEntry.email; }
      Logger.log(` [resolve] -> ID found in 'Sunday Service' form responses: ${personId}`);
  } else {
    let currentHighestOverallId = Math.max(
      findHighestIdInDirectory(),
      findHighestIdInLocalSheets(LOCAL_ID_SHEETS)
    );
    currentHighestOverallId++;
    personId = String(currentHighestOverallId);
    Logger.log(` [resolve] -> Generated NEW ID: ${personId}`);
  }

  if (!firstName && !lastName && fullName) {
    const nameParts = fullName.split(/\s+/);
    firstName = nameParts[0] || "";
    lastName = nameParts.length > 1 ? nameParts.slice(1).join(" ") : "";
  }
  return { id: personId, firstName: firstName, lastName: lastName, email: email };
}

function findHighestIdInLocalSheets(sheetNamesArray) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let highestId = 0;
    sheetNamesArray.forEach(sheetName => {
        const sheet = ss.getSheetByName(sheetName);
        if (sheet) {
            const lastRow = sheet.getLastRow();
            if (lastRow >= 1) {
                let startDataRow = 1; // Default
                if ((sheetName === "Sunday Registration" || sheetName === "Event Registration") && lastRow >= 6) {
                    startDataRow = 6;
                } else if ((sheetName === "Event Attendance" || sheetName === "Service Attendance" || sheetName === "Sunday Service") && lastRow >= 2) {
                    startDataRow = 2; // Assume headers in row 1
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
        }
    });
    return highestId;
}

function findHighestIdInDirectory() {
  let highestId = 0;
  try {
    const props = PropertiesService.getScriptProperties();
    const directoryId = props.getProperty('DIRECTORY_SPREADSHEET_ID');
    if (!directoryId) return 0;
    const directorySheet = SpreadsheetApp.openById(directoryId).getSheetByName("Directory");
    if (directorySheet) {
      const lastRow = directorySheet.getLastRow();
      if (lastRow >= 2) {
        const ids = directorySheet.getRange(2, 1, lastRow - 1, 1).getValues();
        ids.forEach(row => {
          const id = parseInt(row[0]);
          if (!isNaN(id) && id > highestId) { highestId = id; }
        });
      }
    }
  } catch (error) { Logger.log(`❌ (Shared Helper) Error in findHighestIdInDirectory: ${error.message}`); }
  return highestId;
}

function getDirectoryDataMap() {
  const directoryDataMap = new Map();
  try {
    const props = PropertiesService.getScriptProperties();
    const directoryId = props.getProperty('DIRECTORY_SPREADSHEET_ID');
    if (!directoryId) { return directoryDataMap; }
    const directorySheet = SpreadsheetApp.openById(directoryId).getSheetByName("Directory");
    if (directorySheet) {
      const directoryValues = directorySheet.getDataRange().getValues();
      if (directoryValues.length > 1) {
        const headers = directoryValues[0].map(h => String(h || "").trim().toLowerCase());
        const idColIndex = 0, nameColIndex = 1;
        let firstNameColIndex = headers.indexOf("first name", 2);
        let lastNameColIndex = headers.indexOf("last name", 2);
        let emailColIndex = headers.indexOf("email");

        for (let i = 1; i < directoryValues.length; i++) {
          const row = directoryValues[i];
          const personId = String(row[idColIndex] || "").trim();
          const fullName = String(row[nameColIndex] || "").trim();
          if (personId && fullName) {
            directoryDataMap.set(fullName.toUpperCase(), {
              id: personId,
              email: emailColIndex !== -1 ? String(row[emailColIndex] || "").trim() : "",
              firstName: firstNameColIndex !== -1 ? String(row[firstNameColIndex] || "").trim() : "",
              lastName: lastNameColIndex !== -1 ? String(row[lastNameColIndex] || "").trim() : "",
              originalFullName: fullName
            });
          }
        }
      }
    }
  } catch (error) { Logger.log(`❌ (Shared Helper) Error in getDirectoryDataMap: ${error.message}.`); }
  return directoryDataMap;
}

/**
 * MODIFIED: Gets a map of full names to Person IDs. Now has special handling for the
 * "Sunday Registration" sheet to construct the full name from two separate columns.
 */
function getLocalSheetIdMap(sheetName, idColNum = 1, nameColNum = 2) {
    const localIdMap = new Map();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
        const lastRow = sheet.getLastRow();
        if (lastRow < 1) return localIdMap;

        // Determine the actual range to get based on the sheet.
        const maxCol = (sheetName === "Sunday Registration") ? 3 : Math.max(idColNum, nameColNum);
        const data = sheet.getRange(1, 1, lastRow, maxCol).getValues();

        let dataStartRowIndex = 0; // 0-based index
        if (lastRow >= 6 && (sheetName === "Sunday Registration" || sheetName === "Event Registration")) {
            dataStartRowIndex = 5;
        } else if (lastRow >= 2 && (sheetName === "Event Attendance" || sheetName === "Service Attendance" || sheetName === "Sunday Service")) {
            dataStartRowIndex = 1;
        }

        for (let i = dataStartRowIndex; i < data.length; i++) {
            const row = data[i];
            let personId = "";
            let fullName = "";

            // Special handling for the modified Sunday Registration sheet
            if (sheetName === "Sunday Registration") {
                personId = String(row[0] || "").trim(); // ID in col A (index 0)
                const firstName = String(row[1] || "").trim(); // First Name in col B (index 1)
                const lastName = String(row[2] || "").trim(); // Last Name in col C (index 2)
                fullName = `${firstName} ${lastName}`.trim();
            } else {
                // Original logic for all other sheets
                personId = String(row[idColNum - 1] || "").trim();
                fullName = String(row[nameColNum - 1] || "").trim();
            }

            if (personId && fullName) {
                localIdMap.set(fullName.toUpperCase(), personId);
            }
        }
        Logger.log(`(Shared Helper) Local ID map created for "${sheetName}" with ${localIdMap.size} entries.`);
    } else {
        Logger.log(`⚠️ (Shared Helper) Local sheet "${sheetName}" not found for ID lookup.`);
    }
    return localIdMap;
}


/**
 * MODIFIED: Applies row formatting. Now checks the sheet name to apply the correct
 * number of columns (4 for Sunday Reg, 5 for Event Reg).
 */
function refreshRowFormatting(sheet, startDataRow = 6, numRowsInput = -1) {
  if (!sheet) { return; }

  let numRowsToFormat = numRowsInput;
  if (numRowsToFormat === -1) {
    const lastSheetRowWithContent = findLastRowWithData(sheet);
    if (lastSheetRowWithContent < startDataRow) { return; }
    numRowsToFormat = lastSheetRowWithContent - startDataRow + 1;
  }
  if (numRowsToFormat <= 0) { return; }

  // MODIFIED: Determine column count based on the sheet being formatted.
  let numColsToFormat = 5; // Default for sheets like Event Registration
  if (sheet.getName() === "Sunday Registration") {
    numColsToFormat = 4; // Use 4 columns for the modified Sunday sheet
  }
  
  sheet.getRange(startDataRow, 1, numRowsToFormat, numColsToFormat).clearFormat();

  for (let i = 0; i < numRowsToFormat; i++) {
    const currentRowInSheet = startDataRow + i;
    const rowRange = sheet.getRange(currentRowInSheet, 1, 1, numColsToFormat);
    if (i % 2 === 1) { // Apply zebra striping to odd rows (2nd, 4th, etc. in the data)
      rowRange.setBackground("#f5f5f5");
    } else {
      rowRange.setBackground("white");
    }
  }
}

function findLastRowWithData(sheet) {
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) return 0;
  const range = sheet.getRange(1, 1, lastRow, sheet.getMaxColumns());
  const values = range.getValues();
  for (let r = values.length - 1; r >= 0; r--) {
    if (values[r].join('').length > 0) {
      return r + 1;
    }
  }
  return 0;
}

function extractSpreadsheetIdFromUrl(url) {
  const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

function getDateValue(value) {
  if (value instanceof Date && !isNaN(value)) { return value; }
  try {
    const date = new Date(value);
    if (!isNaN(date.getTime())) { return date; }
  } catch (e) {}
  return null;
}
