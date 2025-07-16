/**
 * Event Registration System (Standalone Project - Final Version)
 * Manages event registration, attendance tracking, and attendance statistics.
 * This version checks Directory, Attendance Stats, Service Attendance, and Event Attendance for ID mapping.
 * It relies on the 'DIRECTORY_SPREADSHEET_ID' being set in its own Script Properties.
 * All shared helper functions are duplicated in this project to ensure self-containment.
 */

// --- Event Registration Functions ---

/**
 * Creates or recreates the main "Event Registration" sheet.
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
  sheet.getRange("A1:D1").merge().setHorizontalAlignment("center");
  sheet.getRange("A2").setValue("📅 Event Date:");
  sheet.getRange("B2").setValue(new Date()).setNumberFormat("MM/dd/yyyy");
  sheet.getRange("A3").setValue("📝 Instructions: Check the box next to each person who is present today");
  sheet.getRange("A3:D3").merge();
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


  const headers = ["Person ID", "First Name", "Last Name", "✓ Present"];
  sheet.getRange("A5:D5").setValues([headers]).setFontWeight("bold").setBackground("#4285f4").setFontColor("white");

  sheet.setColumnWidth(1, 100); // Person ID
  sheet.setColumnWidth(2, 150); // First Name
  sheet.setColumnWidth(3, 150); // Last Name
  sheet.setColumnWidth(4, 80);  // Present

  sheet.hideColumns(1); // Hide the ID column, it's for backend tracking

  sheet.getRange("A1:D4").setBackground("#f8f9fa");
  sheet.getRange("A2:B2").setBackground("#e3f2fd");
  sheet.getRange("A4:D4").setBackground("#fff3e0");
  sheet.setFrozenRows(5); // Freeze the header rows
  Logger.log("✅ Event Registration sheet layout created.");
}

/**
 * Populates the "Event Registration" sheet with members from the external Directory,
 * now checking "Attendance Stats" and "Service Attendance" for IDs.
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
      SpreadsheetApp.getUi().alert("Setup Required", "Please set the 'Directory Spreadsheet URL' via the '⚙️ Event Config' menu first.", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
  }
  
  // --- CHANGE: Added "Service Attendance" map back for comprehensive checking ---
  const attendanceStatsIdMap = getLocalSheetIdMap("Attendance Stats", 1, 2);
  const serviceAttendanceIdMap = getLocalSheetIdMap("Service Attendance", 1, 2);
  const eventAttendanceIdMap = getLocalSheetIdMap("Event Attendance", 1, 2);

  let nextGeneratedId = findHighestIdInDirectory();
  if (nextGeneratedId === 0) {
    nextGeneratedId = findHighestIdInLocalSheets(["Event Registration", "Event Attendance", "Sunday Registration", "Attendance Stats", "Service Attendance", "Sunday Service"]); 
  }
  Logger.log(`Initial base for nextGeneratedId: ${nextGeneratedId}`);

  const personsForRegistration = [];
  const processedNewPersonsInThisRun = new Map();

  for (const [normalizedFullName, directoryEntry] of directoryMap.entries()) {
    let personId = directoryEntry.id;
    let firstName = directoryEntry.firstName;
    let lastName = directoryEntry.lastName;
    const fullName = directoryEntry.originalFullName;

    if (!personId) {
      // --- CHANGE: Added "Service Attendance" check back into the logic ---
      const attendanceStatsEntryId = attendanceStatsIdMap.get(normalizedFullName);
      const serviceEntryId = serviceAttendanceIdMap.get(normalizedFullName);
      const eventEntryId = eventAttendanceIdMap.get(normalizedFullName);
      const alreadyProcessedNew = processedNewPersonsInThisRun.get(normalizedFullName);

      if (attendanceStatsEntryId) {
        personId = attendanceStatsEntryId;
      } else if (serviceEntryId) {
        personId = serviceEntryId;
      } else if (eventEntryId) {
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

    personsForRegistration.push([personId, firstName, lastName, false]);
  }

  personsForRegistration.sort((a, b) => {
    const lastNameA = (String(a[2]) || "").toLowerCase(); // Index 2 is Last Name
    const lastNameB = (String(b[2]) || "").toLowerCase();
    const firstNameA = (String(a[1]) || "").toLowerCase(); // Index 1 is First Name
    const firstNameB = (String(b[1]) || "").toLowerCase();

    if (lastNameA < lastNameB) return -1;
    if (lastNameA > lastNameB) return 1;
    if (firstNameA < firstNameB) return -1;
    if (firstNameA > firstNameB) return 1;
    return 0;
  });

  const lastDataRowOnSheet = regSheet.getLastRow();
  if (lastDataRowOnSheet > 5) {
    regSheet.getRange(6, 1, lastDataRowOnSheet - 5, 4).clearContent().clearFormat();
  }
  if (personsForRegistration.length > 0) {
    const startRow = 6;
    regSheet.getRange(startRow, 1, personsForRegistration.length, 4).setValues(personsForRegistration);
    const checkboxRange = regSheet.getRange(startRow, 4, personsForRegistration.length, 1);
    checkboxRange.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
    regSheet.getRange(startRow, 1, personsForRegistration.length, 4).setBorder(true, true, true, true, true, true);
    refreshRowFormatting(regSheet, startRow, personsForRegistration.length);
  }
  regSheet.getRange("D4").setValue(`Status: ${personsForRegistration.length} members loaded`);
  Logger.log(`✅ Registration list populated with ${personsForRegistration.length} members.`);
  updateEventAttendanceCounts(regSheet);
}

/**
 * Allows a user to manually add a new person to the "Event Registration" list.
 */
function addPersonToEventRegistration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Event Registration");
  const sundayRegSheet = ss.getSheetByName("Sunday Registration");

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
    const existingNames = regSheet.getRange(6, 2, lastDataRow - 5, 2).getValues(); // Get First and Last names
    if (existingNames.some(row => (row[0] + " " + row[1]).trim().toLowerCase() === fullNameEntered.toLowerCase())) {
      ui.alert('Duplicate Entry', 'This person is already in the current Event registration list.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
  }

  let personIdToAdd;
  let firstNameToAdd = "";
  let lastNameToAdd = "";
  const normalizedFullName = fullNameEntered.toUpperCase();

  const directoryMap = getDirectoryDataMap();
  // --- CHANGE: Added "Service Attendance" map back for comprehensive checking ---
  const attendanceStatsIdMap = getLocalSheetIdMap("Attendance Stats", 1, 2);
  const serviceAttendanceIdMap = getLocalSheetIdMap("Service Attendance", 1, 2);
  const eventAttendanceIdMap = getLocalSheetIdMap("Event Attendance", 1, 2);

  let sundayRegSheetId = null;
  if (sundayRegSheet) {
    const lastRowSundayReg = sundayRegSheet.getLastRow();
    if (lastRowSundayReg >= 6) {
        const sundayRegSheetData = sundayRegSheet.getRange(6, 1, lastRowSundayReg - 5, 2).getValues(); // Col A=ID, B=Full Name
        for (const row of sundayRegSheetData) {
            if (String(row[1] || "").trim().toUpperCase() === normalizedFullName) {
                sundayRegSheetId = String(row[0] || "").trim();
                break;
            }
        }
    }
  }

  const directoryEntry = directoryMap.get(normalizedFullName);
  // --- CHANGE: Added "Service Attendance" check back into the logic ---
  const attendanceStatsEntryId = attendanceStatsIdMap.get(normalizedFullName);
  const serviceEntryId = serviceAttendanceIdMap.get(normalizedFullName);
  const eventEntryId = eventAttendanceIdMap.get(normalizedFullName);

  if (directoryEntry && directoryEntry.id) {
    personIdToAdd = directoryEntry.id;
    firstNameToAdd = directoryEntry.firstName;
    lastNameToAdd = directoryEntry.lastName;
  } else if (attendanceStatsEntryId) {
    personIdToAdd = attendanceStatsEntryId;
  } else if (serviceEntryId) {
    personIdToAdd = serviceEntryId;
  } else if (eventEntryId) {
    personIdToAdd = eventEntryId;
  } else if (sundayRegSheetId) {
    personIdToAdd = sundayRegSheetId;
  } else {
    let currentHighestOverallId = Math.max(
      findHighestIdInDirectory(),
      findHighestIdInLocalSheets(["Event Registration", "Event Attendance", "Sunday Registration", "Attendance Stats", "Service Attendance", "Sunday Service"])
    );
    currentHighestOverallId++;
    personIdToAdd = String(currentHighestOverallId);
  }

  if (!firstNameToAdd && fullNameEntered) {
    const nameParts = fullNameEntered.split(/\s+/);
    firstNameToAdd = nameParts[0] || "";
    lastNameToAdd = nameParts.length > 1 ? nameParts.slice(1).join(" ") : "";
  }

  const nextSheetRow = (lastDataRow < 5) ? 6 : lastDataRow + 1;
  const newRowData = [personIdToAdd, firstNameToAdd, lastNameToAdd, false];
  regSheet.getRange(nextSheetRow, 1, 1, 4).setValues([newRowData]);
  regSheet.getRange(nextSheetRow, 4).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  const newRowRange = regSheet.getRange(nextSheetRow, 1, 1, 4);
  newRowRange.setBorder(true, true, true, true, true, true);
  refreshRowFormatting(regSheet);

  ui.alert('Person Added!', `${fullNameEntered} has been added with ID ${personIdToAdd}.`, ui.ButtonSet.OK);
  Logger.log(`✅ Manually added ${fullNameEntered} (ID: ${personIdToAdd}) to Event registration list.`);
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
  const eventId = "";

  const spreadsheetTimezone = ss.getSpreadsheetTimeZone();
  const formattedEventDate = Utilities.formatDate(eventDateValue, spreadsheetTimezone, "MM/dd/yyyy");

  regSheet.getRange("D4").setValue("Status: Processing...");

  try {
    const attendanceSheet = ss.getSheetByName("Event Attendance");
    if (!attendanceSheet) {
      regSheet.getRange("D4").setValue("Status: Error - Event Attendance sheet missing");
      throw new Error("'Event Attendance' sheet not found");
    }

    const directoryMap = getDirectoryDataMap();
    const lastRegDataRow = regSheet.getLastRow();
    if (lastRegDataRow < 6) {
      regSheet.getRange("D4").setValue("Status: No members to process");
      return;
    }

    const regData = regSheet.getRange(6, 1, lastRegDataRow - 5, 4).getValues();
    const attendanceEntries = [];
    let checkedCount = 0;

    for (const row of regData) {
      const [personId, firstName, lastName, isChecked] = row;
      const fullName = `${firstName || ''} ${lastName || ''}`.trim();

      if (isChecked === true && fullName) {
        let email = "";
        const normalizedFullName = fullName.toUpperCase();
        const directoryEntry = directoryMap.get(normalizedFullName);
        if (directoryEntry) {
          email = directoryEntry.email || "";
        }

        // Structure for 'Event Attendance': A=Person ID, B=Full Name, C=Event Name, ..., K=Event Date, ...
        attendanceEntries.push([
          personId, fullName, eventName, eventId, firstName || "", lastName || "",
          email, "", "", "", formattedEventDate, '', '', new Date()
        ]);
        checkedCount++;
      }
    }

    if (attendanceEntries.length === 0) {
      regSheet.getRange("D4").setValue("Status: No members checked");
      return;
    }

    if (attendanceSheet.getLastRow() < 1) {
      const attendanceHeaders = ["Person ID", "Full Name", "Event Name", "Event ID", "First Name", "Last Name", "Email", "Phone Number", "Form Sheet", "Role", "Event Date", "First Time?", "Needs Follow-up?", "Timestamp"];
      attendanceSheet.getRange(1, 1, 1, attendanceHeaders.length).setValues([attendanceHeaders]).setFontWeight("bold").setBackground("#4285f4").setFontColor("white");
    }

    const nextRowAttendanceSheet = findLastRowWithData(attendanceSheet) + 1;
    attendanceSheet.getRange(nextRowAttendanceSheet, 1, attendanceEntries.length, attendanceEntries[0].length).setValues(attendanceEntries);
    attendanceSheet.getRange(nextRowAttendanceSheet, 11, attendanceEntries.length, 1).setNumberFormat("MM/dd/yyyy");
    attendanceSheet.getRange(nextRowAttendanceSheet, 14, attendanceEntries.length, 1).setNumberFormat("MM/dd/yyyy HH:mm:ss");

    regSheet.getRange(6, 4, lastRegDataRow - 5, 1).setValue(false);
    regSheet.getRange("D4").setValue(`Status: ${checkedCount} attendees submitted`);
    SpreadsheetApp.getUi().alert( 'Attendance Submitted!', `Successfully submitted attendance for ${checkedCount} members.`);
    
    updateEventAttendanceCounts(regSheet);
    processEventAttendanceForFollowUpByName();

  } catch (error) {
    regSheet.getRange("D4").setValue("Status: Error occurred");
    Logger.log(`❌ Error submitting attendance: ${error.message}\n${error.stack || ""}`);
    SpreadsheetApp.getUi().alert("Error", `Error submitting attendance: ${error.message}`);
  }
}

/**
 * Clears all checkboxes in the "Event Registration" sheet.
 */
function clearAllEventChecks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Event Registration");
  if (!regSheet) return;
  
  const lastDataRow = regSheet.getLastRow();
  if (lastDataRow >= 6) {
    regSheet.getRange(6, 4, lastDataRow - 5, 1).setValue(false); // Checkbox is in column D (4)
    regSheet.getRange("D4").setValue("Status: All checks cleared");
  } else {
    regSheet.getRange("D4").setValue("Status: No checks to clear");
  }
}

/**
 * Adds or re-applies checkboxes to the 'Present' column (Column D).
 */
function addCheckboxesToEventRegistration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Event Registration");
  if (!regSheet) return;

  const lastRow = regSheet.getLastRow();
  if (lastRow < 6) return;
  
  // Checks for a name in column B (First Name) to determine how many rows to format
  const nameValues = regSheet.getRange(6, 2, lastRow - 5, 1).getValues();
  let rowsWithActualNames = 0;
  for (let i = 0; i < nameValues.length; i++) {
    if (String(nameValues[i][0] || "").trim() !== "") {
      rowsWithActualNames = i + 1;
    }
  }
  if (rowsWithActualNames === 0) return;

  try {
    const checkboxRange = regSheet.getRange(6, 4, rowsWithActualNames, 1); // Checkbox is in column D (4)
    checkboxRange.clearContent().setValue(false).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());

    const dataFormattingRange = regSheet.getRange(6, 1, rowsWithActualNames, 4);
    dataFormattingRange.setBorder(true, true, true, true, true, true);
    refreshRowFormatting(regSheet, 6, rowsWithActualNames);

    regSheet.getRange("D4").setValue(`Status: ${rowsWithActualNames} members ready`);
    SpreadsheetApp.getUi().alert('Checkboxes Added/Reformatted!', `Successfully added/reformatted checkboxes for ${rowsWithActualNames} attendee rows.`);
  } catch (error) {
    Logger.log(`❌ Error adding/reformatting checkboxes: ${error.message}`);
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
  updateEventAttendanceCounts(regSheet);
  Logger.log("✅ Empty Event Registration sheet created.");
  SpreadsheetApp.getUi().alert(
    'Empty Registration Sheet Created!',
    'Event Registration sheet is ready for manual data entry.\n\n' +
    'INSTRUCTIONS:\n' +
    '1. Paste attendee data starting row 6 (First Name in Col B, Last in C).\n' +
    '2. Use "📋 Event Check-in" → "🔲 Add/Reformat Checkboxes" to set up column D.\n' +
    '3. Enter event date in B2 and start checking attendance!',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Removes a person from the "Event Registration" list by full name.
 */
function removePersonFromEventRegistration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Event Registration");
  if (!regSheet) return;

  const ui = SpreadsheetApp.getUi();
  const lastDataRow = regSheet.getLastRow();
  if (lastDataRow < 6) {
    ui.alert('No People', 'No people in the list to remove.', ui.ButtonSet.OK);
    return;
  }

  const nameResponse = ui.prompt('Remove Person', 'Enter the FULL NAME of the person to remove (case-insensitive):', ui.ButtonSet.OK_CANCEL);
  if (nameResponse.getSelectedButton() !== ui.Button.OK) return;
  const nameToRemove = String(nameResponse.getResponseText() || "").trim().toLowerCase();
  if (!nameToRemove) return;

  const allData = regSheet.getRange(6, 1, lastDataRow - 5, 3).getValues(); // Get ID, First, Last
  let rowToDeleteInSheet = -1;
  for (let i = 0; i < allData.length; i++) {
    const sheetFullName = `${allData[i][1]} ${allData[i][2]}`.trim().toLowerCase(); // Combine First and Last Name
    if (sheetFullName === nameToRemove) {
      rowToDeleteInSheet = i + 6;
      break;
    }
  }

  if (rowToDeleteInSheet > 0) {
    regSheet.deleteRow(rowToDeleteInSheet);
    ui.alert('Person Removed!', `'${nameResponse.getResponseText().trim()}' has been removed.`, ui.ButtonSet.OK);
    refreshRowFormatting(regSheet);
    updateEventAttendanceCounts(regSheet);
  } else {
    ui.alert('Not Found', `Person '${nameResponse.getResponseText().trim()}' not found.`);
  }
}

/**
 * Sorts the data in the "Event Registration" sheet by Last Name (Column C).
 */
function sortEventRegistrationByLastName() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Event Registration");
  if (!regSheet) return;

  const lastDataRow = findLastRowWithData(regSheet);
  if (lastDataRow < 6) return;

  const numDataRows = lastDataRow - 5;
  if (numDataRows <= 0) return;

  const dataRange = regSheet.getRange(6, 1, numDataRows, 4); // Range is now 4 columns
  // Sort by Last Name (Col C, which is column #3), then First Name (Col B, which is column #2)
  dataRange.sort([{ column: 3, ascending: true }, { column: 2, ascending: true }]);
  
  refreshRowFormatting(regSheet, 6, numDataRows);
  SpreadsheetApp.getUi().alert("Event list sorted by Last Name.");
}

/**
 * Adds a custom menu to the spreadsheet UI for Event Registration functions.
 * NOTE: This function would typically be called by an onOpen() trigger.
 */
function addEventRegistrationMenu() {
  SpreadsheetApp.getUi().createMenu('📋 Event Check-in')
    .addItem('📁 Get Names from Directory', 'populateEventRegistrationList')
    .addItem('✅ Submit Attendance', 'submitEventRegistrationAttendance')
    .addSeparator()
    .addItem('➕ Add Attendee (Quick Add)', 'addPersonToEventRegistration')
    .addItem('➖ Remove Attendee', 'removePersonFromEventRegistration')
    .addItem('🔲 Add/Reformat Checkboxes', 'addCheckboxesToEventRegistration')
    .addItem('Sort by Last Name', 'sortEventRegistrationByLastName')
    .addSeparator()
    .addItem('🆕 Create Empty Registration Sheet', 'createEmptyEventRegistrationSheet')
    .addItem('🔍 Update Follow-up Flags', 'processEventAttendanceForFollowUpByName')
    .addToUi();
}

/**
 * Updates the attendance counters on the Event Registration sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} regSheet The Event Registration sheet.
 */
function updateEventAttendanceCounts(regSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attendanceSheet = ss.getSheetByName("Event Attendance");
  if (!attendanceSheet) {
    regSheet.getRange("G2:H2").setValue("N/A");
    return;
  }

  const attendanceData = attendanceSheet.getDataRange().getValues();
  if (attendanceData.length < 2) {
    regSheet.getRange("G2:H2").setValue(0);
    return;
  }

  const currentEventName = regSheet.getRange("A1").getValue().replace(/^🏛️\s*/, "").trim();
  const currentMonth = new Date().getMonth();
  const currentYear = new Date().getFullYear();

  let totalAttendeesMonthCount = 0;
  const uniqueEventAttendancesThisEvent = new Set();

  const PERSON_ID_COL_IDX = 0;   // Col A
  const EVENT_NAME_COL_IDX = 2;  // Col C
  const EVENT_DATE_COL_IDX = 10; // Col K

  for (let i = 1; i < attendanceData.length; i++) {
    const row = attendanceData[i];
    const personId = String(row[PERSON_ID_COL_IDX] || "").trim();
    const eventNameInRecord = String(row[EVENT_NAME_COL_IDX] || "").trim();
    const eventDateInRecord = getDateValue(row[EVENT_DATE_COL_IDX]);

    if (!personId || !eventDateInRecord) continue;

    if (eventDateInRecord.getMonth() === currentMonth && eventDateInRecord.getFullYear() === currentYear) {
      totalAttendeesMonthCount++;
    }

    if (eventNameInRecord.toLowerCase() === currentEventName.toLowerCase()) {
      const compositeKeyForEventAttendance = `${personId}_${Utilities.formatDate(eventDateInRecord, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd")}`;
      uniqueEventAttendancesThisEvent.add(compositeKeyForEventAttendance);
    }
  }

  regSheet.getRange("G2").setValue(totalAttendeesMonthCount);
  regSheet.getRange("H2").setValue(uniqueEventAttendancesThisEvent.size);
}

// --- Shared Helper Functions ---

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
                } else if ((sheetName.includes("Attendance") || sheetName.includes("Service") || sheetName === "Attendance Stats") && lastRow >= 2) {
                    startDataRow = 2;
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
    const directoryId = PropertiesService.getScriptProperties().getProperty('DIRECTORY_SPREADSHEET_ID');
    if (!directoryId) return 0;
    
    const directorySheet = SpreadsheetApp.openById(directoryId).getSheetByName("Directory");
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
    }
  } catch (error) {
    Logger.log(`❌ Error in findHighestIdInDirectory: ${error.message}`);
  }
  return highestId;
}

function getDirectoryDataMap() {
  const directoryDataMap = new Map();
  try {
    const directoryId = PropertiesService.getScriptProperties().getProperty('DIRECTORY_SPREADSHEET_ID');
    if (!directoryId) return directoryDataMap;
    
    const directorySheet = SpreadsheetApp.openById(directoryId).getSheetByName("Directory");
    if (directorySheet) {
      const directoryValues = directorySheet.getDataRange().getValues();
      if (directoryValues.length > 1) {
        const headers = directoryValues[0].map(h => String(h || "").trim().toLowerCase());
        const idColIndex = 0;
        const nameColIndex = 1;
        let firstNameColIndex = headers.indexOf("first name");
        if (firstNameColIndex === -1) firstNameColIndex = 2;
        let lastNameColIndex = headers.indexOf("last name");
        if (lastNameColIndex === -1) lastNameColIndex = 3;
        let emailColIndex = headers.indexOf("email");
        if (emailColIndex === -1) emailColIndex = 7;

        for (let i = 1; i < directoryValues.length; i++) {
          const row = directoryValues[i];
          const personId = String(row[idColIndex] || "").trim();
          const fullName = String(row[nameColIndex] || "").trim();
          if (personId && fullName) {
            const normalizedFullName = fullName.toUpperCase();
            directoryDataMap.set(normalizedFullName, {
              id: personId,
              email: row[emailColIndex] ? String(row[emailColIndex]).trim() : "",
              firstName: row[firstNameColIndex] ? String(row[firstNameColIndex]).trim() : "",
              lastName: row[lastNameColIndex] ? String(row[lastNameColIndex]).trim() : "",
              originalFullName: fullName
            });
          }
        }
      }
    }
  } catch (error) {
    Logger.log(`❌ Error in getDirectoryDataMap: ${error.message}`);
  }
  return directoryDataMap;
}

function getLocalSheetIdMap(sheetName, idColNum = 1, nameColNum = 2) {
    const localIdMap = new Map();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (sheet) {
        const lastRow = sheet.getLastRow();
        if (lastRow >= 1) {
            const data = sheet.getRange(1, 1, lastRow, Math.max(idColNum, nameColNum)).getValues();
            let dataStartRowIndex = 0; // Default to 0 for header-less sheets or unknown formats

            if (lastRow >= 6 && (sheetName === "Sunday Registration" || sheetName === "Event Registration")) {
                dataStartRowIndex = 5; // Start reading data from row 6 (index 5)
            } else if (lastRow >= 2 && (sheetName.includes("Attendance") || sheetName.includes("Service") || sheetName === "Attendance Stats")) {
                dataStartRowIndex = 1; // Start reading data from row 2 (index 1)
            }
            
            for (let i = dataStartRowIndex; i < data.length; i++) {
                const personId = String(data[i][idColNum - 1] || "").trim();
                const fullName = String(data[i][nameColNum - 1] || "").trim();
                if (personId && fullName) {
                    localIdMap.set(fullName.toUpperCase(), personId);
                }
            }
        }
    }
    return localIdMap;
}

function refreshRowFormatting(sheet, startDataRow = 6, numRowsInput = -1) {
  if (!sheet) return;

  let numRowsToFormat = numRowsInput;
  if (numRowsToFormat === -1) {
    const lastSheetRowWithContent = findLastRowWithData(sheet);
    if (lastSheetRowWithContent < startDataRow) return;
    numRowsToFormat = lastSheetRowWithContent - startDataRow + 1;
  }

  if (numRowsToFormat <= 0) return;

  sheet.getRange(startDataRow, 1, numRowsToFormat, 4).clearFormat(); // Use 4 columns

  for (let i = 0; i < numRowsToFormat; i++) {
    const currentRowInSheet = startDataRow + i;
    const rowRange = sheet.getRange(currentRowInSheet, 1, 1, 4); // Use 4 columns
    if (i % 2 === 1) { // Alternate row color
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
  if (value instanceof Date && !isNaN(value)) {
    return value;
  }
  if (typeof value === 'string' || typeof value === 'number') {
    try {
      const date = new Date(value);
      if (!isNaN(date.getTime())) {
        return date;
      }
    } catch (e) { /* ignore parse error */ }
  }
  return null;
}
