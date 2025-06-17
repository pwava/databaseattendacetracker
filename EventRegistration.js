/**
 * Event Registration System (Standalone Project - Final Version)
 * Manages event registration, attendance tracking, and attendance statistics.
 * This script is intended to be its own Apps Script project, separate from any Sunday Service script.
 * It relies on the 'DIRECTORY_SPREADSHEET_ID' being set in its own Script Properties,
 * which should point to the same central Directory as the Sunday script.
 *
 * All shared helper functions are duplicated in this project to ensure self-containment and avoid
 * "not defined" errors when the Sunday project is not active or correctly bound.
 */

// --- Event Registration Functions ---

/**
 * Creates or recreates the main "Event Registration" sheet.
 * Prompts the user if the sheet already exists.
 */
function createEventRegistrationSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let regSheet = ss.getSheetByName("Event Registration");
  if (regSheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Sheet Already Exists',
      'Event Registration sheet already exists. Do you want to recreate it? This will clear all existing data.',
      ui.ButtonSet.YES_NO
    );
    if (response === ui.Button.YES) {
      ss.deleteSheet(regSheet);
    } else {
      return; // User chose not to recreate
    }
  }
  regSheet = ss.insertSheet("Event Registration");
  setupEventRegistrationSheetLayout(regSheet);
  populateEventRegistrationList(regSheet);
  // Update counts immediately after creation/population
  updateEventAttendanceCounts(regSheet);
  Logger.log("✅ Event Registration sheet created successfully! Person IDs are populated via new logic.");
  SpreadsheetApp.getUi().alert(
    'Registration Sheet Created!',
    'Event Registration sheet has been created and populated with active members.\n\n' +
    'Person IDs in Column A are fetched/generated based on Directory, local sheets, or new.\n\n' +
    'The registration team can now:\n' +
    '1. Enter the event date in cell B2\n' +
    '2. Change the event name in cell A1 (e.g., "Kids Camp Registration")\n' +
    '3. Check the boxes for attendees\n' +
    '4. Click "Submit Attendance" from the "📋 Event Check-in" menu to transfer to Event Attendance sheet\n\n' +
    'Menus have been added/updated for easy access to functions.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Sets up the initial layout, headers, and basic formatting for the Event registration sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to format.
 */
function setupEventRegistrationSheetLayout(sheet) {
  sheet.clear();
  sheet.getRange("A1").setValue("🏛️ Custom Event Name REGISTRATION").setFontSize(16).setFontWeight("bold");
  sheet.getRange("A1:E1").merge().setHorizontalAlignment("center");
  sheet.getRange("A2").setValue("📅 Event Date:");
  sheet.getRange("B2").setValue(new Date()).setNumberFormat("MM/dd/yyyy");
  sheet.getRange("A3").setValue("📝 Instructions: Check the box next to each person who is present today");
  sheet.getRange("A3:E3").merge();
  sheet.getRange("A4").setValue("🔄 Refresh List");
  sheet.getRange("B4").setValue("✅ Submit Attendance");
  sheet.getRange("C4").setValue("🧹 Clear All Checks");
  sheet.getRange("D4").setValue("Status: Ready");

  // Add the attendance counters labels
  sheet.getRange("G1").setValue("Attendees Total This Month").setFontWeight("bold");
  sheet.getRange("H1").setValue("Attendees this event").setFontWeight("bold");
  sheet.getRange("G1:H1").setBackground("#e3f2fd").setHorizontalAlignment("center");
  // G2 and H2 will be populated by script, not formulas
  sheet.getRange("G2:H2").setBackground("#e3f2fd").setHorizontalAlignment("center");


  const headers = ["Person ID", "Full Name", "First Name", "Last Name", "✓ Present"];
  sheet.getRange("A5:E5").setValues([headers]).setFontWeight("bold").setBackground("#4285f4").setFontColor("white");

  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 140);
  sheet.setColumnWidth(5, 80);

  sheet.hideColumns(1); // Hide the ID column, it's for backend tracking

  sheet.getRange("A1:E4").setBackground("#f8f9fa");
  sheet.getRange("A2:B2").setBackground("#e3f2fd");
  sheet.getRange("A4:D4").setBackground("#fff3e0");
  sheet.setFrozenRows(5); // Freeze the header rows
  Logger.log("✅ Event Registration sheet layout created.");
}

/**
 * Populates the "Event Registration" sheet with members from the external Directory,
 * assigning existing IDs or generating new ones.
 */
function populateEventRegistrationList(regSheet = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!regSheet) {
    regSheet = ss.getSheetByName("Event Registration");
    if (!regSheet) {
      Logger.log("❌ Event Registration sheet not found for populateEventRegistrationList");
      SpreadsheetApp.getUi().alert("Error", "Event Registration sheet not found. Please create it first.", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
  }

  const directoryMap = getDirectoryDataMap();
  if (directoryMap.size === 0 && PropertiesService.getScriptProperties().getProperty('DIRECTORY_SPREADSHEET_ID')) {
    Logger.log("⚠️ Directory map is empty but DIRECTORY_SPREADSHEET_ID is set. Check Directory sheet content or ID validity.");
    SpreadsheetApp.getUi().alert("Warning", "Could not load data from Directory. Please ensure the Directory Spreadsheet URL is set correctly and the 'Directory' sheet contains data.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  } else if (directoryMap.size === 0 && !PropertiesService.getScriptProperties().getProperty('DIRECTORY_SPREADSHEET_ID')) {
      SpreadsheetApp.getUi().alert("Setup Required", "Please set the 'Directory Spreadsheet URL' via the '⚙️ Event Config' menu first.", SpreadsheetApp.getUi().ButtonSet.OK); // Updated menu name
      return;
  }

  const eventAttendanceIdMap = getLocalSheetIdMap("Event Attendance", 1, 2);

  let nextGeneratedId = findHighestIdInDirectory();
  if (nextGeneratedId === 0) {
    // Explicitly list local sheets relevant to THIS project for highest ID search
    // This project will also look at Sunday/Service Attendance if those sheets are visible from this project.
    nextGeneratedId = findHighestIdInLocalSheets(["Event Registration", "Event Attendance", "Sunday Registration", "Service Attendance", "Sunday Service"]); 
  }
  Logger.log(`Initial base for nextGeneratedId (starting with Directory, then local to this project): ${nextGeneratedId}`);

  const personsForRegistration = [];
  const processedNewPersonsInThisRun = new Map();

  for (const [normalizedFullName, directoryEntry] of directoryMap.entries()) {
    let personId = directoryEntry.id;
    let firstName = directoryEntry.firstName;
    let lastName = directoryEntry.lastName;
    const fullName = directoryEntry.originalFullName;

    if (!personId) {
      const eventEntryId = eventAttendanceIdMap.get(normalizedFullName);
      const alreadyProcessedNew = processedNewPersonsInThisRun.get(normalizedFullName);

      if (eventEntryId) {
        personId = eventEntryId;
      } else if (alreadyProcessedNew) {
        personId = alreadyProcessedNew.id;
        firstName = alreadyProcessedNew.firstName || firstName;
        lastName = alreadyProcessedNew.lastName || lastName;
      } else {
        nextGeneratedId++;
        personId = String(nextGeneratedId);
        processedNewPersonsInThisRun.set(normalizedFullName, { id: personId, firstName: firstName, lastName: lastName });
        Logger.log(`Generated new ID ${personId} for ${fullName} (from Directory but ID missing).`);
      }
    }

    if (!firstName && !lastName && fullName) {
      const nameParts = fullName.split(/\s+/);
      firstName = nameParts[0] || "";
      lastName = nameParts.length > 1 ? nameParts.slice(1).join(" ") : "";
    }

    personsForRegistration.push([personId, fullName, firstName, lastName, false]);
  }

  personsForRegistration.sort((a, b) => {
    const lastNameA = (String(a[3]) || "").toLowerCase();
    const lastNameB = (String(b[3]) || "").toLowerCase();
    const firstNameA = (String(a[2]) || "").toLowerCase();
    const firstNameB = (String(b[2]) || "").toLowerCase();

    if (lastNameA < lastNameB) return -1;
    if (lastNameA > lastNameB) return 1;
    if (firstNameA < firstNameB) return -1;
    if (firstNameA > firstNameB) return 1;
    return 0;
  });

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
  Logger.log(`✅ Registration list populated with ${personsForRegistration.length} members. IDs fetched/generated.`);

  // Update event attendance counts after populating the list
  updateEventAttendanceCounts(regSheet);
}

/**
 * Allows a user to manually add a new person to the "Event Registration" list.
 * This function now includes a check against the Sunday Registration sheet
 * and a robust ID lookup/assignment process.
 */
function addPersonToEventRegistration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Event Registration");
  const sundayRegSheet = ss.getSheetByName("Sunday Registration"); // Get Sunday Registration sheet for cross-check/ID reuse

  if (!regSheet) {
    SpreadsheetApp.getUi().alert("Error", "Event Registration sheet not found", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  const ui = SpreadsheetApp.getUi();

  const nameResponse = ui.prompt('Add Person', 'Enter the full name:', ui.ButtonSet.OK_CANCEL);
  if (nameResponse.getSelectedButton() !== ui.Button.OK) return;
  const fullNameEntered = String(nameResponse.getResponseText() || "").trim();
  if (!fullNameEntered) {
    ui.alert('Input Error', 'Please enter a valid name', ui.ButtonSet.OK);
    return;
  }

  // Check for duplicate on the CURRENT Event registration sheet first
  const lastDataRow = regSheet.getLastRow();
  if (lastDataRow >= 6) {
    const existingNamesOnCurrentSheet = regSheet.getRange(6, 2, lastDataRow - 5, 1).getValues();
    if (existingNamesOnCurrentSheet.some(row => row[0] && String(row[0]).trim().toLowerCase() === fullNameEntered.toLowerCase())) {
      ui.alert('Duplicate Entry', 'This person is already in the current Event registration list. No need to add again.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
  }

  let personIdToAdd;
  let firstNameToAdd = "";
  let lastNameToAdd = "";
  const normalizedFullName = fullNameEntered.toUpperCase();

  // Fetch all relevant ID maps
  const directoryMap = getDirectoryDataMap();
  const serviceAttendanceIdMap = getLocalSheetIdMap("Service Attendance", 1, 2); // Get Service Attendance map
  const eventAttendanceIdMap = getLocalSheetIdMap("Event Attendance", 1, 2); // Get Event Attendance map

  // Check the Sunday Registration sheet for an ID, to reuse it.
  let sundayRegSheetId = null;
  if (sundayRegSheet) { // Check if the Sunday Registration sheet exists
    const lastRowSundayReg = sundayRegSheet.getLastRow();
    if (lastRowSundayReg >= 6) {
      const sundayRegSheetData = sundayRegSheet.getRange(6, 1, lastRowSundayReg - 5, 2).getValues(); // Get ID (Col A) and Full Name (Col B)
      for (const row of sundayRegSheetData) {
        if (String(row[1] || "").trim().toUpperCase() === normalizedFullName) {
          sundayRegSheetId = String(row[0] || "").trim();
          Logger.log(`Found ID ${sundayRegSheetId} for ${fullNameEntered} in the 'Sunday Registration' sheet.`);
          break;
        }
      }
    }
  }

  // --- Robust ID Lookup Order (Prioritize existing IDs from all sources) ---
  const directoryEntry = directoryMap.get(normalizedFullName);
  const serviceEntryId = serviceAttendanceIdMap.get(normalizedFullName);
  const eventEntryId = eventAttendanceIdMap.get(normalizedFullName);


  if (directoryEntry && directoryEntry.id) {
    personIdToAdd = directoryEntry.id;
    firstNameToAdd = directoryEntry.firstName;
    lastNameToAdd = directoryEntry.lastName;
    Logger.log(`Found ID ${personIdToAdd} for ${fullNameEntered} in Directory.`);
  } else if (serviceEntryId) {
    personIdToAdd = serviceEntryId;
    if (directoryEntry) {
      firstNameToAdd = directoryEntry.firstName || firstNameToAdd;
      lastNameToAdd = directoryEntry.lastName || lastNameToAdd;
    }
    Logger.log(`Found ID ${personIdToAdd} for ${fullNameEntered} in Service Attendance.`);
  } else if (eventEntryId) {
    personIdToAdd = eventEntryId;
    if (directoryEntry) {
      firstNameToAdd = directoryEntry.firstName || firstNameToAdd;
      lastNameToAdd = directoryEntry.lastName || lastNameToAdd;
    }
    Logger.log(`Found ID ${personIdToAdd} for ${fullNameEntered} in Event Attendance.`);
  } else if (sundayRegSheetId) { // Use ID from Sunday Registration sheet if available
      personIdToAdd = sundayRegSheetId;
      if (directoryEntry) {
        firstNameToAdd = directoryEntry.firstName || firstNameToAdd;
        lastNameToAdd = directoryEntry.lastName || lastNameToAdd;
      }
      Logger.log(`Found ID ${personIdToAdd} for ${fullNameEntered} in the 'Sunday Registration' sheet. Reusing this ID.`);
  }
  else {
    // If not found in any existing source, generate a new ID
    // Explicitly list all potential ID sources across both projects for finding the highest ID
    let currentHighestOverallId = Math.max(
      findHighestIdInDirectory(),
      findHighestIdInLocalSheets(["Event Registration", "Event Attendance", "Sunday Registration", "Service Attendance", "Sunday Service"]) // Explicitly list all potential ID sources for max ID
    );
    currentHighestOverallId++;
    personIdToAdd = String(currentHighestOverallId);
    Logger.log(`Generated new ID ${personIdToAdd} for manually added (Event) ${fullNameEntered} (not found in any existing source).`);
  }
  // --- END Robust ID LOGIC ---

  if (!firstNameToAdd && fullNameEntered) {
    const nameParts = fullNameEntered.split(/\s+/);
    firstNameToAdd = nameParts[0] || "";
    lastNameToAdd = nameParts.length > 1 ? nameParts.slice(1).join(" ") : "";
  }

  const nextSheetRow = (lastDataRow < 5) ? 6 : lastDataRow + 1;
  const newRowData = [personIdToAdd, fullNameEntered, firstNameToAdd, lastNameToAdd, false];
  regSheet.getRange(nextSheetRow, 1, 1, 5).setValues([newRowData]);
  regSheet.getRange(nextSheetRow, 5).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  const newRowRange = regSheet.getRange(nextSheetRow, 1, 1, 5);
  newRowRange.setBorder(true, true, true, true, true, true);
  refreshRowFormatting(regSheet);

  ui.alert('Person Added!', `${fullNameEntered} has been added with ID ${personIdToAdd}.`, ui.ButtonSet.OK);
  Logger.log(`✅ Manually added ${fullNameEntered} (ID: ${personIdToAdd}) to Event registration list.`);
  // Update event attendance counts after adding a person
  updateEventAttendanceCounts(regSheet);
}


/**
 * Submits checked-in attendees from "Event Registration" to the "Event Attendance" sheet.
 */
function submitEventRegistrationAttendance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Event Registration");
  if (!regSheet) {
    SpreadsheetApp.getUi().alert("Error", "Event Registration sheet not found", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const eventDateValue = regSheet.getRange("B2").getValue();
  if (!eventDateValue || !(eventDateValue instanceof Date) || isNaN(eventDateValue.getTime())) {
    SpreadsheetApp.getUi().alert("Input Error", "Please enter a valid event date in cell B2.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const eventName = regSheet.getRange("A1").getValue().replace(/^🏛️\s*/, "").trim();
  const eventId = ""; // Placeholder for Event ID

  const spreadsheetTimezone = ss.getSpreadsheetTimeZone();
  const formattedEventDate = Utilities.formatDate(eventDateValue, spreadsheetTimezone, "MM/dd/yyyy");

  regSheet.getRange("D4").setValue("Status: Processing...");

  try {
    const attendanceSheet = ss.getSheetByName("Event Attendance");
    if (!attendanceSheet) {
      SpreadsheetApp.getUi().alert("Error", "'Event Attendance' sheet not found. Please ensure it exists.", SpreadsheetApp.getUi().ButtonSet.OK);
      regSheet.getRange("D4").setValue("Status: Error - Event Attendance sheet missing");
      throw new Error("'Event Attendance' sheet not found");
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
        let phoneNumber = "";
        let formSheet = "";
        let role = "";

        const normalizedFullName = String(fullName).trim().toUpperCase();
        const directoryEntry = directoryMap.get(normalizedFullName);
        if (directoryEntry) {
          email = directoryEntry.email || "";
        } else {
          Logger.log(`Email/other details not found in Directory for ${fullName} (ID: ${personId}). Will submit blank.`);
        }

        // Structure for 'Event Attendance' sheet:
        // A=Person ID, B=Full Name, C=Event Name, D=Event ID, E=First Name, F=Last Name,
        // G=Email, H=Phone Number, I=Form Sheet, J=Role, K=Event Date, L=First Time?, M=Needs Follow-up?, N=Timestamp
        attendanceEntries.push([
          personId,
          fullName,
          eventName,          // C
          eventId,            // D
          firstName || "",    // E
          lastName || "",     // F
          email,              // G
          phoneNumber,        // H
          formSheet,          // I
          role,               // J
          formattedEventDate, // K (Event Date)
          '',                 // L (First Time? - calculated later by processEventAttendanceForFollowUpByName)
          '',                 // M (Needs Follow-up? - calculated later by processEventAttendanceForFollowUpByName)
          new Date()          // N (Timestamp)
        ]);
        checkedCount++;
      }
    }

    if (attendanceEntries.length === 0) {
      regSheet.getRange("D4").setValue("Status: No members checked");
      SpreadsheetApp.getUi().alert("No Checks", "No members were checked for attendance.", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // Ensure attendance sheet has headers if new/empty
    if (attendanceSheet.getLastRow() < 1) {
      const attendanceHeaders = ["Person ID", "Full Name", "Event Name", "Event ID", "First Name", "Last Name", "Email", "Phone Number", "Form Sheet", "Role", "Event Date", "First Time?", "Needs Follow-up?", "Timestamp"];
      attendanceSheet.getRange(1, 1, 1, attendanceHeaders.length).setValues([attendanceHeaders]).setFontWeight("bold").setBackground("#4285f4").setFontColor("white");
    }

    const nextRowAttendanceSheet = findLastRowWithData(attendanceSheet) + 1;
    attendanceSheet.getRange(nextRowAttendanceSheet, 1, attendanceEntries.length, attendanceEntries[0].length).setValues(attendanceEntries);
    // Format Event Date column (K)
    attendanceSheet.getRange(nextRowAttendanceSheet, 11, attendanceEntries.length, 1).setNumberFormat("MM/dd/yyyy");
    // Format Timestamp column (N)
    attendanceSheet.getRange(nextRowAttendanceSheet, 14, attendanceEntries.length, 1).setNumberFormat("MM/dd/yyyy HH:mm:ss");

    regSheet.getRange(6, 5, lastRegDataRow - 5, 1).setValue(false);
    regSheet.getRange("D4").setValue(`Status: ${checkedCount} attendees submitted`);
    SpreadsheetApp.getUi().alert(
      'Attendance Submitted!',
      `Successfully submitted attendance for ${checkedCount} members to 'Event Attendance' sheet.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    Logger.log(`✅ Successfully submitted ${checkedCount} attendance entries.`);

    // Update event attendance counts after submission
    updateEventAttendanceCounts(regSheet);
    // Process for follow-up flags after submission
    processEventAttendanceForFollowUpByName(); // Call the new function here

  } catch (error) {
    regSheet.getRange("D4").setValue("Status: Error occurred");
    Logger.log(`❌ Error submitting attendance: ${error.message}\n${error.stack || ""}`);
    SpreadsheetApp.getUi().alert("Error", `Error submitting attendance: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Clears all checkboxes in the "Event Registration" sheet.
 */
function clearAllEventChecks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Event Registration");
  if (!regSheet) {
    SpreadsheetApp.getUi().alert("Error", "Event Registration sheet not found", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  const lastDataRow = regSheet.getLastRow();
  if (lastDataRow >= 6) {
    regSheet.getRange(6, 5, lastDataRow - 5, 1).setValue(false);
    regSheet.getRange("D4").setValue("Status: All checks cleared");
    Logger.log("✅ All Event checkboxes cleared");
  } else {
    regSheet.getRange("D4").setValue("Status: No checks to clear");
    Logger.log("ℹ️ No data rows found to clear Event checks from.");
  }
}

/**
 * Adds or re-applies checkboxes to the 'Present' column (Column E)
 * of the "Event Registration" sheet.
 */
function addCheckboxesToEventRegistration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Event Registration");
  if (!regSheet) {
    SpreadsheetApp.getUi().alert("Error", "Event Registration sheet not found", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  const lastRow = regSheet.getLastRow();
  if (lastRow < 6) {
    SpreadsheetApp.getUi().alert("No Data", "No data found below row 5 to add checkboxes to. Please add attendee data starting row 6.", SpreadsheetApp.getUi().ButtonSet.OK);
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
    SpreadsheetApp.getUi().alert('Checkboxes Added/Reformatted!', `Successfully added/reformatted checkboxes for ${rowsWithActualNames} attendee rows.\n\nSheet is ready:\n1. Enter event date in B2\n2. Check attendance\n3. Click Submit Attendance`, SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log(`✅ Added/Reformatted checkboxes to ${rowsWithActualNames} rows for Event Registration`);
  } catch (error) {
    Logger.log(`❌ Error adding/reformatting checkboxes for Event Registration: ${error.message}`);
    SpreadsheetApp.getUi().alert("Error", `Error adding/reformatting checkboxes for Event Registration: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Creates an empty "Event Registration" sheet.
 */
function createEmptyEventRegistrationSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let regSheet = ss.getSheetByName("Event Registration");
  if (regSheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert('Sheet Already Exists', 'Event Registration sheet already exists. Recreate it as empty? This will clear all existing data.', ui.ButtonSet.YES_NO);
    if (response === ui.Button.YES) ss.deleteSheet(regSheet);
    else return;
  }
  regSheet = ss.insertSheet("Event Registration");
  setupEventRegistrationSheetLayout(regSheet);
  // Update counts on empty sheet setup
  updateEventAttendanceCounts(regSheet);
  Logger.log("✅ Empty Event Registration sheet created.");
  SpreadsheetApp.getUi().alert(
    'Empty Registration Sheet Created!',
    'Event Registration sheet is ready for manual data entry (Columns A-E for data, starting row 6).\n\n' +
    'Person IDs (Col A) will be fetched from Directory or generated if you use Refresh/Add Attendee.\n\n' +
    'INSTRUCTIONS:\n' +
    '1. Change "Custom Event Name" in cell A1 to your actual event name.\n' +
    '2. Paste directory data starting row 6 (Full Name in Col B, First in C, Last in D - ID will be handled by other functions)\n' +
    '3. Use "📋 Event Check-in" → "🔲 Add/Reformat Checkboxes" to set up column E.\n' +
    '4. Enter event date in B2 and start checking attendance!',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Removes a person from the "Event Registration" list by full name.
 */
function removePersonFromEventRegistration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Event Registration");
  if (!regSheet) {
    SpreadsheetApp.getUi().alert("Error", "Event Registration sheet not found", SpreadsheetApp.getUi().ButtonSet.OK);
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
    ui.alert('Input Error', 'Please enter a valid name', ui.ButtonSet.OK);
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
    Logger.log(`✅ Removed '${nameResponse.getResponseText().trim()}' from Event registration list, row ${rowToDeleteInSheet}`);
    refreshRowFormatting(regSheet);
    // Update event attendance counts after removing a person
    updateEventAttendanceCounts(regSheet);
  } else {
    ui.alert('Not Found', `Person '${nameResponse.getResponseText().trim()}' not found in the Event registration list.`, ui.ButtonSet.OK);
  }
}

/**
 * Sorts the data in the "Event Registration" sheet by Last Name (Column D).
 */
function sortEventRegistrationByLastName() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Event Registration");
  if (!regSheet) {
    SpreadsheetApp.getUi().alert("Error", "Event Registration sheet not found.", SpreadsheetApp.getUi().ButtonSet.OK);
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
  dataRange.sort([{ column: 4, ascending: true }, { column: 3, ascending: true }]);
  refreshRowFormatting(regSheet, 6, numDataRows);
  SpreadsheetApp.getUi().alert("Event list sorted by Last Name.");
  Logger.log("✅ Event Registration list sorted by last name.");
}

/**
 * Adds a custom menu to the spreadsheet UI for Event Registration functions.
 */
function addEventRegistrationMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📋 Event Check-in')
    .addItem('📁 Get Names from Directory', 'populateEventRegistrationList')
    .addItem('✅ Submit Attendance', 'submitEventRegistrationAttendance')
    .addSeparator()
    .addItem('➕ Add Attendee (Quick Add)', 'addPersonToEventRegistration')
    .addItem('🔲 Add/Reformat Checkboxes', 'addCheckboxesToEventRegistration')
    .addItem('Sort by Last Name', 'sortEventRegistrationByLastName')
    .addSeparator()
    .addItem('🆕 Create Empty Registration Sheet', 'createEmptyEventRegistrationSheet')
    .addItem('🔍 Update Follow-up Flags', 'processEventAttendanceForFollowUpByName')
    .addToUi();
  Logger.log("✅ Event Check-in menu definition attempted by addEventRegistrationMenu.");
}

/**
 * Updates the 'Attendees Total' and 'Attendees this event' counters on the Event Registration sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} regSheet The Event Registration sheet.
 */
function updateEventAttendanceCounts(regSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attendanceSheet = ss.getSheetByName("Event Attendance");
  if (!attendanceSheet) {
    regSheet.getRange("G2").setValue("N/A");
    regSheet.getRange("H2").setValue(0); // Set to 0 if sheet is missing
    Logger.log("Event Attendance sheet not found for updating counters.");
    return;
  }

  const attendanceData = attendanceSheet.getDataRange().getValues();
  if (attendanceData.length < 2) { // No data rows, only headers
    regSheet.getRange("G2").setValue(0);
    regSheet.getRange("H2").setValue(0);
    Logger.log("No data in Event Attendance sheet, counters set to 0.");
    return;
  }

  const currentEventName = regSheet.getRange("A1").getValue().replace(/^🏛️\s*/, "").trim();
  const currentMonth = new Date().getMonth(); // 0-indexed month (0 = Jan, 11 = Dec)
  const currentYear = new Date().getFullYear();

  let totalAttendeesMonthCount = 0; // Counts all entries within the month
  const uniqueEventAttendancesThisEvent = new Set(); // Stores composite keys (Person ID + Event Date) for the current event

  // Updated column indices based on new Event Attendance layout
  const PERSON_ID_COL_IDX = 0; // Column A
  const EVENT_NAME_COL_IDX = 2; // Column C
  const EVENT_DATE_COL_IDX = 10; // Column K (Event Date) - This is the date from the Event Registration sheet

  for (let i = 1; i < attendanceData.length; i++) { // Start from 1 to skip headers
    const row = attendanceData[i];
    const personId = String(row[PERSON_ID_COL_IDX] || "").trim();
    const eventNameInRecord = String(row[EVENT_NAME_COL_IDX] || "").trim();
    const eventDateInRecord = getDateValue(row[EVENT_DATE_COL_IDX]); // Use helper to get valid date object

    if (!personId || !eventDateInRecord) continue; // Skip rows without a Person ID or valid Event Date

    // Logic for 'Attendees Total This Month' (G2) - count ALL entries whose Event Date falls in the current month
    if (eventDateInRecord.getMonth() === currentMonth &&
        eventDateInRecord.getFullYear() === currentYear) {
      totalAttendeesMonthCount++; // Increment for each entry whose event date falls in the current month
    }

    // Logic for 'Attendees this event' (H2) - unique (Person ID + Event Date) for the CURRENT Event Name
    // This creates a composite key: "ID_YYYY-MM-DD" based on EVENT DATE
    if (eventNameInRecord.toLowerCase() === currentEventName.toLowerCase()) {
      const compositeKeyForEventAttendance = `${personId}_${Utilities.formatDate(eventDateInRecord, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd")}`;
      uniqueEventAttendancesThisEvent.add(compositeKeyForEventAttendance);
    }
  }

  regSheet.getRange("G2").setValue(totalAttendeesMonthCount); // Count of all event participations in the current month
  regSheet.getRange("H2").setValue(uniqueEventAttendancesThisEvent.size); // Count of unique (Person ID, Event Date) for the current event name
  Logger.log(`Event Attendance Counters Updated: Attendees Total This Month: ${totalAttendeesMonthCount}, Attendees this event (unique person-event-date combos for current event): ${uniqueEventAttendancesThisEvent.size}`);
}

// --- Shared Helper Functions (Defined in both projects for full independence) ---

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
                // These conditions now specifically look for headers based on known sheets
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
            // Log this as info, not error, since other project's sheets might not exist yet
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
    } catch (e) {
      // Fall through if parsing fails
    }
  }
  // For numbers that represent dates (e.g., from CSV import without formatting)
  if (typeof value === 'number' && value > 0 && value < 2958466) {
      try {
          const date = new Date((value - 25569) * 86400 * 1000); // Convert Excel/Sheets date serial to milliseconds
          if (!isNaN(date.getTime())) return date;
      } catch(e) {}
  }
  return null;
}