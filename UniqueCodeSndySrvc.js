/**
 * Main handler for form submissions to 'sunday service'.
 * It first ensures names are combined, then processes IDs, then transfers data.
 * @param {Object} e The event object from an onFormSubmit trigger.
 */
function processSundayServiceIds(e) {
  const functionStartTime = new Date();
  Logger.log('processSundayServiceIds script started.');

  // --- Configuration ---
  // DIRECTORY_SPREADSHEET_ID will now be fetched from Script Properties via getDirectorySpreadsheetIdFromProperties()
  const SUNDAY_SERVICE_SHEET_NAME = 'sunday service';
  const EVENT_ATTENDANCE_SHEET_NAME = 'Event attendance'; // Used for highest ID and map
  const SUNDAY_REGISTRATION_SHEET_NAME = 'sunday registration'; // ADDED FOR ID MAPPING
  const EVENT_REGISTRATION_SHEET_NAME = 'event registration';   // ADDED FOR ID MAPPING
  const DIRECTORY_TAB_NAME = 'directory';       // Used for highest ID and map
  const NEW_MEMBER_FORM_TAB_NAME = 'new member form';     // Used for highest ID and map
  const SERVICE_ATTENDANCE_SHEET_NAME = 'Service Attendance'; // Used for mapping
  const ID_COLUMN = 1;        // Personal ID is in Column A
  const NAME_COLUMN = 2;        // Full Name is in Column B (this is what combineNamesOnFormSubmit populates)
  const MAX_ROWS_TO_PROCESS_MANUALLY = 1500; // Max rows for manual run

  // --- Determine Mode (Trigger or Manual) ---
  const isTriggerMode = (e && e.range);

  if (isTriggerMode) {
    Logger.log(`Running in ON FORM SUBMIT trigger mode for row: ${e.range.getRow()}.`);

    // ****** Call combineNamesOnFormSubmit first in trigger mode ******
    try {
      if (typeof combineNamesOnFormSubmit === "function") {
        Logger.log(`Attempting to call combineNamesOnFormSubmit(e) for row ${e.range.getRow()} to prepare Full Name...`);
        combineNamesOnFormSubmit(e); // Call it, passing the event object
        SpreadsheetApp.flush(); // IMPORTANT: Ensure changes from combineNames are saved before proceeding
        Logger.log('combineNamesOnFormSubmit(e) finished. Full Name should now be populated in Column B.');
      } else {
        Logger.log('CRITICAL WARNING: combineNamesOnFormSubmit function not found. Full Name may not be prepared in Column B.');
      }
    } catch (combineError) {
      Logger.log(`ERROR during explicit call to combineNamesOnFormSubmit: ${combineError.toString()}`);
      Logger.log(`Proceeding with ID assignment, but Full Name in Column B for row ${e.range.getRow()} might be missing or stale due to this error.`);
    }
    // ****** END NEW SECTION ******

  } else {
    Logger.log('Running in MANUAL or non-event (e.g., time-driven) trigger mode.');
  }

  // --- Get Spreadsheet Objects ---
  let currentSs;
  try {
    currentSs = SpreadsheetApp.getActiveSpreadsheet();
  } catch (err) {
    Logger.log(`Error accessing active spreadsheet: ${err.toString()}`);
    return;
  }

  let directorySs;
  let DIRECTORY_SPREADSHEET_ID; // Declare it here to be used in this function scope
  try {
    // Fetch DIRECTORY_SPREADSHEET_ID from Script Properties
    DIRECTORY_SPREADSHEET_ID = getDirectorySpreadsheetIdFromProperties();
    directorySs = SpreadsheetApp.openById(DIRECTORY_SPREADSHEET_ID);
  } catch (err) {
    Logger.log(`Error opening directory spreadsheet (ID: ${DIRECTORY_SPREADSHEET_ID || 'not set'}): ${err.toString()}. Some ID data might be incomplete or property not set.`);
    directorySs = null; // Allow script to continue but be aware some data might be missing
  }

  Logger.log('Fetching data from all source sheets for ID processing...');
  const sundayServiceDataAll = getSheetData(currentSs, SUNDAY_SERVICE_SHEET_NAME);
  const eventAttendanceData = getSheetData(currentSs, EVENT_ATTENDANCE_SHEET_NAME);
  const sundayRegistrationData = getSheetData(currentSs, SUNDAY_REGISTRATION_SHEET_NAME);
  const eventRegistrationData = getSheetData(currentSs, EVENT_REGISTRATION_SHEET_NAME);
  const directoryData = directorySs ? getSheetData(directorySs, DIRECTORY_TAB_NAME) : [];
  const newMemberFormData = directorySs ? getSheetData(directorySs, NEW_MEMBER_FORM_TAB_NAME) : [];
  const serviceAttendanceData = getSheetData(currentSs, SERVICE_ATTENDANCE_SHEET_NAME);

  if (sundayServiceDataAll.length === 0 && isTriggerMode) {
    Logger.log(`"${SUNDAY_SERVICE_SHEET_NAME}" sheet is critically empty or could not be read, but in trigger mode. Will attempt to process the specific event row.`);
  } else if (sundayServiceDataAll.length === 0 && !isTriggerMode) {
    Logger.log(`"${SUNDAY_SERVICE_SHEET_NAME}" sheet is critically empty or could not be read. Script cannot proceed in manual mode.`);
    return;
  }

  Logger.log('Calculating highest existing numeric ID...');
  let highestExistingNumber = 0;
  
  function updateHighestNumberFromSheetData(data, sheetNameForLog) {
    if (!data || data.length === 0) {
      Logger.log(`No data from "${sheetNameForLog}" to scan for highest ID.`);
      return;
    }
    for (let i = 1; i < data.length; i++) { // Assumes header in row 0
      if (data[i] && data[i].length >= ID_COLUMN) {
        const idCell = data[i][ID_COLUMN - 1];
        if (idCell !== null && idCell !== undefined && String(idCell).trim() !== "") {
          const number = extractNumberFromId(String(idCell)); 
          if (!isNaN(number)) {
            highestExistingNumber = Math.max(highestExistingNumber, number);
          }
        }
      }
    }
    Logger.log(`After scanning "${sheetNameForLog}", current max ID number is: ${highestExistingNumber}`);
  }

  updateHighestNumberFromSheetData(sundayServiceDataAll, SUNDAY_SERVICE_SHEET_NAME);
  updateHighestNumberFromSheetData(eventAttendanceData, EVENT_ATTENDANCE_SHEET_NAME);
  updateHighestNumberFromSheetData(sundayRegistrationData, SUNDAY_REGISTRATION_SHEET_NAME);
  updateHighestNumberFromSheetData(eventRegistrationData, EVENT_REGISTRATION_SHEET_NAME);
  updateHighestNumberFromSheetData(newMemberFormData, NEW_MEMBER_FORM_TAB_NAME);
  updateHighestNumberFromSheetData(directoryData, DIRECTORY_TAB_NAME);
  updateHighestNumberFromSheetData(serviceAttendanceData, SERVICE_ATTENDANCE_SHEET_NAME);
  Logger.log(`Initial highest existing numeric ID across all sheets: ${highestExistingNumber}`);

  Logger.log('Building master Name-ID lookup map...');
  const masterNameIdMap = new Map();
  
  function populateMapFromSheetData(data, sheetNameForLog) {
    if (!data || data.length === 0) {
      Logger.log(`No data from "${sheetNameForLog}" to populate map.`);
      return;
    }
    for (let i = 1; i < data.length; i++) { // Assumes header
      if (data[i] && data[i].length >= NAME_COLUMN && data[i].length >= ID_COLUMN) {
        const nameCell = data[i][NAME_COLUMN - 1];
        const idCell = data[i][ID_COLUMN - 1];
        if (nameCell && String(nameCell).trim() !== "" && idCell !== null && idCell !== undefined && String(idCell).trim() !== "") {
          masterNameIdMap.set(String(nameCell).trim().toUpperCase(), String(idCell).trim());
        }
      }
    }
    Logger.log(`Populated map from "${sheetNameForLog}". Map size now: ${masterNameIdMap.size}`);
  }
  
  populateMapFromSheetData(sundayServiceDataAll, SUNDAY_SERVICE_SHEET_NAME);
  populateMapFromSheetData(eventAttendanceData, EVENT_ATTENDANCE_SHEET_NAME);
  populateMapFromSheetData(sundayRegistrationData, SUNDAY_REGISTRATION_SHEET_NAME);
  populateMapFromSheetData(eventRegistrationData, EVENT_REGISTRATION_SHEET_NAME);
  populateMapFromSheetData(newMemberFormData, NEW_MEMBER_FORM_TAB_NAME);
  populateMapFromSheetData(directoryData, DIRECTORY_TAB_NAME);
  populateMapFromSheetData(serviceAttendanceData, SERVICE_ATTENDANCE_SHEET_NAME);
  Logger.log(`Master Name-ID map built. Total unique names with IDs: ${masterNameIdMap.size}`);

  const sundayServiceSheet = currentSs.getSheetByName(SUNDAY_SERVICE_SHEET_NAME);
  if (!sundayServiceSheet) {
    Logger.log(`CRITICAL Error: Could not get sheet: "${SUNDAY_SERVICE_SHEET_NAME}". Exiting ID processing.`);
    return;
  }
  
  let rowsToProcessDetails = [];
  if (isTriggerMode) {
    const firstSubmittedSheetRow = e.range.getRow();
    const numSubmittedRows = e.range.getNumRows();
    Logger.log(`Trigger event detail: sheet row(s) ${firstSubmittedSheetRow} to ${firstSubmittedSheetRow + numSubmittedRows - 1}.`);
    for (let i = 0; i < numSubmittedRows; i++) {
      const currentSheetRow = firstSubmittedSheetRow + i;
      if (currentSheetRow > 0) {
        rowsToProcessDetails.push({ sheetRow: currentSheetRow });
      } else {
        Logger.log(`Warning: Submitted sheet row ${currentSheetRow} is invalid during trigger mode. Skipping.`);
      }
    }
    if (rowsToProcessDetails.length === 0 && numSubmittedRows > 0) {
      Logger.log("Warning: No valid rows from trigger event to process, though event had range. Exiting ID assignment.");
      return;
    }
  } else { // Manual or non-event Time-driven Mode
    const actualEndDataIndex = Math.min(sundayServiceDataAll.length, MAX_ROWS_TO_PROCESS_MANUALLY + 1); // +1 to account for header if MAX_ROWS is for data rows
    if (actualEndDataIndex <= 1 && sundayServiceDataAll.length <=1 ) {
      Logger.log(`"${SUNDAY_SERVICE_SHEET_NAME}" sheet has no data rows to process (or only header) in manual/time-driven mode. Script finished.`);
      return;
    }
    Logger.log(`Manual/Time-driven Mode: Processing 'sunday service' from sheet row 2 up to ${actualEndDataIndex}.`);
    for (let dataIdx = 1; dataIdx < actualEndDataIndex; dataIdx++) { // dataIdx is for sundayServiceDataAll (0-indexed), so dataIdx=1 is sheet row 2
      rowsToProcessDetails.push({ dataIndexInArray: dataIdx, sheetRow: dataIdx + 1 });
    }
  }

  const updatesToWrite = [];
  let idAssignedInTriggerMode = false;

  for (const detail of rowsToProcessDetails) {
    const actualSheetRowNumber = detail.sheetRow;
    let currentRowData;

    if (isTriggerMode) {
      try {
        // Fetch fresh data for the specific row. This is crucial.
        // Ensure enough columns are read, at least up to NAME_COLUMN.
        const numColsToFetch = Math.max(ID_COLUMN, NAME_COLUMN, sundayServiceSheet.getLastColumn() > 0 ? sundayServiceSheet.getLastColumn() : NAME_COLUMN);
        currentRowData = sundayServiceSheet.getRange(actualSheetRowNumber, 1, 1, numColsToFetch).getValues()[0];
        Logger.log(`Trigger Mode: Fetched data for row ${actualSheetRowNumber}: [${currentRowData.join(", ")}]`);
      } catch (fetchErr) {
        Logger.log(`Trigger Mode: Error fetching data for row ${actualSheetRowNumber}: ${fetchErr.toString()}. Skipping this row.`);
        continue;
      }
    } else { // Manual mode - use pre-fetched data from sundayServiceDataAll
      if (detail.dataIndexInArray < sundayServiceDataAll.length) {
        currentRowData = sundayServiceDataAll[detail.dataIndexInArray];
      } else {
        Logger.log(`Manual Mode: dataIndexInArray ${detail.dataIndexInArray} out of bounds for sundayServiceDataAll. Skipping row ${actualSheetRowNumber}.`);
        continue;
      }
    }
    
    if (!currentRowData) {
      Logger.log(`Skipping Sheet Row ${actualSheetRowNumber}: Row data is undefined or could not be fetched.`);
      continue;
    }
    // Check if enough columns exist, especially for NAME_COLUMN
    if (currentRowData.length < NAME_COLUMN) {
      Logger.log(`Skipping Sheet Row ${actualSheetRowNumber}: Row does not have enough columns for Name (needs at least ${NAME_COLUMN}). Has ${currentRowData.length}. Data: [${currentRowData.join(', ')}]`);
      continue;
    }

    const currentName = currentRowData[NAME_COLUMN - 1]; // Array is 0-indexed
    const existingIdInSheet = currentRowData[ID_COLUMN - 1]; // Array is 0-indexed

    if (currentName && String(currentName).trim() !== "") {
      const formattedName = String(currentName).trim().toUpperCase();
      let determinedId = "";

      if (masterNameIdMap.has(formattedName)) {
        determinedId = masterNameIdMap.get(formattedName);
        Logger.log(`Sheet Row ${actualSheetRowNumber}: Name "${currentName}" found in master map. ID from map: "${determinedId}". Existing ID in sheet: "${existingIdInSheet}"`);
      } else {
        highestExistingNumber++;
        determinedId = String(highestExistingNumber);
        Logger.log(`Sheet Row ${actualSheetRowNumber}: Name "${currentName}" NOT found in map. Generating new ID: "${determinedId}". Existing ID in sheet: "${existingIdInSheet}"`);
        masterNameIdMap.set(formattedName, determinedId); // Update map for future consistency within this run
      }

      determinedId = String(determinedId).trim();
      const sheetIdToCompare = String(existingIdInSheet || "").trim();

      if (determinedId !== "" && determinedId !== sheetIdToCompare) {
        updatesToWrite.push({ row: actualSheetRowNumber, id: determinedId });
        Logger.log(`Sheet Row ${actualSheetRowNumber}: QUEUED FOR ID UPDATE. New ID: "${determinedId}", Old ID: "${sheetIdToCompare}".`);
        if (isTriggerMode) idAssignedInTriggerMode = true;
      } else if (determinedId === "") {
        Logger.log(`Sheet Row ${actualSheetRowNumber}: SKIPPED ID UPDATE. Determined ID is empty for name "${currentName}".`);
      } else {
        Logger.log(`Sheet Row ${actualSheetRowNumber}: NO ID UPDATE NEEDED. Determined ID ("${determinedId}") matches existing sheet ID ("${sheetIdToCompare}").`);
      }
    } else {
      Logger.log(`Sheet Row ${actualSheetRowNumber}: No name in Column B (or name is blank). Skipping ID assignment.`);
    }
  }

  if (updatesToWrite.length > 0) {
    Logger.log(`Attempting to write ${updatesToWrite.length} ID updates to "${SUNDAY_SERVICE_SHEET_NAME}".`);
    let successCount = 0;
    updatesToWrite.forEach(update => {
      try {
        sundayServiceSheet.getRange(update.row, ID_COLUMN).setValue(update.id);
        successCount++;
      } catch (err) {
        Logger.log(`Error writing ID "${update.id}" to row ${update.row} in "${SUNDAY_SERVICE_SHEET_NAME}": ${err.toString()}`);
      }
    });
    if (successCount > 0) {
      SpreadsheetApp.flush(); // IMPORTANT: Flush changes to the sheet
      Logger.log(`${successCount} of ${updatesToWrite.length} ID updates successfully written and flushed to "${SUNDAY_SERVICE_SHEET_NAME}".`);
    } else {
      Logger.log(`No ID updates were successfully written, though ${updatesToWrite.length} were queued.`);
    }
  } else {
    Logger.log('No ID updates were necessary for the processed rows in "sunday service".');
  }

  const functionEndTime = new Date();
  const duration = (functionEndTime.getTime() - functionStartTime.getTime()) / 1000;
  Logger.log(`processSundayServiceIds script finished. Duration: ${duration} seconds.`);

  // --- Chaining: Call onFormSubmitTransfer if in trigger mode and ID was assigned ---
  if (isTriggerMode && idAssignedInTriggerMode) {
    Logger.log("ID assignment successful in trigger mode. Attempting to call onFormSubmitTransfer...");
    try {
      if (typeof onFormSubmitTransfer === "function") {
        onFormSubmitTransfer(e); // Pass the original event object
        Logger.log("Successfully called onFormSubmitTransfer(e).");
      } else {
        Logger.log("CRITICAL ERROR: onFormSubmitTransfer function not found. Cannot chain call. Make sure it's in the same script project.");
      }
    } catch (transferErr) {
      Logger.log(`Error occurred during chained call to onFormSubmitTransfer: ${transferErr.toString()}`);
    }
  } else if (isTriggerMode && !idAssignedInTriggerMode) {
    Logger.log("Trigger mode, but no new ID was assigned or written (likely due to missing name). Not calling onFormSubmitTransfer.");
  }
}

//------------------------------------------------------------------
// HELPER FUNCTIONS (ensure these are in your script project)
//------------------------------------------------------------------

/**
 * Retrieves the DIRECTORY_SPREADSHEET_ID from Script Properties.
 * This function should be in the same Apps Script project as processSundayServiceIds.
 * Make sure the 'DIRECTORY_SPREADSHEET_ID' property is set in your project's script properties.
 * (e.g., via Project Settings -> Script Properties or a custom menu function)
 * @returns {string} The ID of the directory spreadsheet.
 * @throws {Error} If the DIRECTORY_SPREADSHEET_ID property is not set.
 */
function getDirectorySpreadsheetIdFromProperties() {
  const props = PropertiesService.getScriptProperties();
  const directoryId = props.getProperty('DIRECTORY_SPREADSHEET_ID');
  if (!directoryId) {
    // It's good practice to throw an error if a critical config is missing.
    throw new Error('DIRECTORY_SPREADSHEET_ID script property not set. Please set it in your Apps Script Project Properties or via your Config menu.');
  }
  return directoryId;
}


/**
 * Helper function to get all data from a sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet The spreadsheet object.
 * @param {string} sheetName The name of the sheet.
 * @return {Array<Array>} The sheet data, or an empty array if sheet not found or empty.
 */
function getSheetData(spreadsheet, sheetName) {
  if (!spreadsheet) {
    Logger.log(`getSheetData: Spreadsheet object is null for sheet name "${sheetName}". Returning empty data.`);
    return [];
  }
  try {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`Warning: Sheet "${sheetName}" not found in spreadsheet "${spreadsheet.getName()}". Returning empty data.`);
      return [];
    }
    const lastRow = sheet.getLastRow();
    if (lastRow < 1) {
      Logger.log(`Sheet "${sheetName}" is empty. Returning empty data.`);
      return [];
    }
    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) {
      Logger.log(`Sheet "${sheetName}" has no columns. Returning empty data.`);
      return [];
    }
    return sheet.getRange(1, 1, lastRow, lastCol).getValues();
  } catch (err) {
    Logger.log(`Error reading from sheet "${sheetName}" in "${spreadsheet.getName()}": ${err.toString()}. Returning empty data.`);
    return [];
  }
}

/**
 * Helper function to extract a number from an ID string, typically the trailing number.
 * @param {string} idString The ID string.
 * @return {number} The extracted number, or NaN if not found.
 */
function extractNumberFromId(idString) {
  if (idString === null || idString === undefined) return NaN;
  const str = String(idString).trim();
  if (/^\d+$/.test(str)) { // If the string is purely numeric
    return parseInt(str, 10);
  }
  const match = str.match(/(\d+)$/); // Try to extract trailing number
  if (match && match[1]) {
    return parseInt(match[1], 10);
  }
  return NaN; // Return NaN if no number could be extracted
}

/**
 * Combines First Name (Col C) and Last Name (Col D) into Full Name (Col B)
 * for the row determined by the event object or the last row.
 * @param {Object} e The event object from an onFormSubmit trigger (optional).
 */
function combineNamesOnFormSubmit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Sunday Service");

  if (!sheet) {
    Logger.log("combineNamesOnFormSubmit: Sheet 'Sunday Service' not found. Exiting.");
    return;
  }
  
  var targetRow;
  if (e && e.range) {
    targetRow = e.range.getRow();
    Logger.log(`combineNamesOnFormSubmit: Processing row ${targetRow} from onFormSubmit event.`);
  } else {
    targetRow = sheet.getLastRow();
    Logger.log(`combineNamesOnFormSubmit: No event object, processing last row: ${targetRow}. This should ideally be called with an event.`);
    if (targetRow === 0) {
      Logger.log("combineNamesOnFormSubmit: Sheet is empty, nothing to process.");
      return;
    }
  }

  var firstName = "";
  var lastName = "";
  try {
    firstName = sheet.getRange(targetRow, 3).getValue(); // Column C
    lastName = sheet.getRange(targetRow, 4).getValue();  // Column D
  } catch (err) {
    Logger.log(`combineNamesOnFormSubmit: Error getting first/last name for row ${targetRow}: ${err.toString()}`);
    return;
  }

  if (firstName && typeof firstName.trim === 'function' && lastName && typeof lastName.trim === 'function' && firstName.trim() !== "" && lastName.trim() !== "") {
    var fullName = firstName.trim() + " " + lastName.trim();
    try {
      sheet.getRange(targetRow, 2).setValue(fullName); // Column B
      Logger.log(`combineNamesOnFormSubmit: Successfully set Full Name "${fullName}" in row ${targetRow}, Column B.`);
    } catch (err) {
      Logger.log(`combineNamesOnFormSubmit: Error setting full name for row ${targetRow}: ${err.toString()}`);
    }
  } else {
    let missing = [];
    if (!firstName || (typeof firstName.trim === 'function' && firstName.trim() === "")) missing.push("First Name (Col C)");
    if (!lastName || (typeof lastName.trim === 'function' && lastName.trim() === "")) missing.push("Last Name (Col D)");
    Logger.log(`combineNamesOnFormSubmit: ${missing.join(' and ')} is missing or blank in row ${targetRow}. Full name not combined.`);
  }
}

// You also need the 'onFormSubmitTransfer(e)' function in your script project.
// If you don't have it, you'll need to add it separately based on your requirements.
// function onFormSubmitTransfer(e) { ... }