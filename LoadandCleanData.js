/**
 * Runs when the spreadsheet is opened to add our custom menus.
 * @param {GoogleAppsScript.Events.SheetsOnOpen} e The event object (optional).
 */
function onOpen(e) {
    Logger.log("Master onOpen triggered. AuthMode: " + (e ? e.authMode : 'N/A Event Object'));

    try {
        // Your original Config menu creation (no change needed here, it still works)
        SpreadsheetApp.getUi()
            .createMenu('⚙️ Config') // Changed to include icon for consistency
            .addItem('Set Directory Spreadsheet URL…', 'showDirectoryDialog') // Changed text for clarity
            .addToUi();
        Logger.log("✅ Config menu added by onOpen.");
    } catch (error) {
        Logger.log("Error adding Config menu in onOpen: " + error.message + " Stack: " + error.stack);
    }

    try {
        // Call Sunday Service menu (ensure addSundayRegistrationMenu is defined)
        addSundayRegistrationMenu();
        Logger.log("Call to addSundayRegistrationMenu completed from onOpen.");
    } catch (error) {
        Logger.log("Error during addSundayRegistrationMenu in onOpen: " + error.message + " Stack: " + error.stack);
    }

    try {
        // Call Event Registration menu (ensure addEventRegistrationMenu is defined)
        addEventRegistrationMenu();
        Logger.log("Call to addEventRegistrationMenu completed from onOpen.");
    } catch (error) {
        Logger.log("Error during addEventRegistrationMenu in onOpen: " + error.message + " Stack: " + error.stack);
    }

    // --- START: ADD THIS NEW BLOCK ---
    try {
        // Call the function to add the User Manual menu
        addUserManualMenu(); 
        Logger.log("Call to addUserManualMenu completed from onOpen.");
    } catch (error) {
        Logger.log("Error during addUserManualMenu in onOpen: " + error.message + " Stack: " + error.stack);
    }
    // --- END: ADD THIS NEW BLOCK ---

    // Update event attendance counters if the active sheet is the Event Registration sheet on open
    try {
      const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      if (activeSheet && activeSheet.getName() === "Event Registration") {
        updateEventAttendanceCounts(activeSheet);
        Logger.log("Event attendance counters updated on open for Event Registration sheet.");
      }
    } catch (error) {
      Logger.log("Error updating event attendance counters on open: " + error.message);
    }
}

/**
 * Prompts the user to paste a Sheets URL or ID,
 * extracts the ID, and saves it as a script property.
 */
function showDirectoryDialog() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    'Configure Directory Spreadsheet',
    'Paste the full Google Sheets URL or just the Spreadsheet ID:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const input = resp.getResponseText().trim();
  const id    = extractSpreadsheetId(input);
  
  if (!id) {
    ui.alert('❌ Invalid URL or ID. Please try again.');
    return;
  }
  
  PropertiesService
    .getScriptProperties()
    .setProperty('DIRECTORY_SPREADSHEET_ID', id);
  
  ui.alert('✅ DIRECTORY_SPREADSHEET_ID set to:\n' + id);
}

/**
 * Helpers: pulls an ID out of either
 *   • a /d/URL segment, or
 *   • a bare ID string
 */
function extractSpreadsheetId(input) {
  const urlMatch = input.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (urlMatch && urlMatch[1]) {
    return urlMatch[1];
  }
  // basic sanity check for a bare ID
  if (/^[a-zA-Z0-9-_]+$/.test(input)) {
    return input;
  }
  return null;
}

/**
 * Fetches raw data from the required Google Sheets.
 * Reads "Service Attendance", "Event Attendance", and "Attendance Stats"
 * from the active spreadsheet, and "Directory" from an external spreadsheet.
 * Uses getDataRange() to fetch all data with content.
 * Includes error handling and logging for debugging.
 *
 * @returns {object} An object containing the data arrays:
 *   { sData, eData, dData, statsData },
 * or undefined on failure.
 */
function getDataFromSheets() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  
  // --- get the Directory ID from script props (fail fast if missing) ---
  const directoryId = props.getProperty('DIRECTORY_SPREADSHEET_ID');
  if (!directoryId) {
    throw new Error(
      "🚨 Missing Script Property: 'DIRECTORY_SPREADSHEET_ID'.\n" +
      "Use Config → Set Directory Spreadsheet… to configure it."
    );
  }
  
  // --- Get local sheets ---
  const serviceSheet = ss.getSheetByName("Service Attendance");
  const eventSheet   = ss.getSheetByName("Event Attendance");
  const statsSheet   = ss.getSheetByName("Attendance Stats");
  
  if (!serviceSheet || !eventSheet || !statsSheet) {
    if (!serviceSheet) Logger.log("❌ 'Service Attendance' not found.");
    if (!eventSheet)   Logger.log("❌ 'Event Attendance' not found.");
    if (!statsSheet)   Logger.log("❌ 'Attendance Stats' not found.");
    return;
  }
  
  // --- Load data from local sheets ---
  let sData, eData, statsData;
  try {
    sData     = serviceSheet.getDataRange().getValues();
    eData     = eventSheet.getDataRange().getValues();
    statsData = statsSheet.getDataRange().getValues();
    Logger.log("✅ Local sheets loaded.");
  } catch (err) {
    Logger.log("❌ Error reading local sheets: " + err.message);
    return;
  }
  
  // --- Load Directory from external spreadsheet by ID ---
  let directorySS;
  try {
    directorySS = SpreadsheetApp.openById(directoryId);
    Logger.log("✅ External spreadsheet opened via ID.");
  } catch (err) {
    Logger.log("❌ Could not open external spreadsheet ID=" + directoryId + " : " + err.message);
    return;
  }
  
  const directorySheet = directorySS.getSheetByName("Directory");
  if (!directorySheet) {
    Logger.log("❌ 'Directory' sheet not found in external spreadsheet.");
    return;
  }
  
  let dData;
  try {
    dData = directorySheet.getDataRange().getValues();
    Logger.log("✅ Directory sheet loaded.");
  } catch (err) {
    Logger.log("❌ Error reading Directory sheet: " + err.message);
    return;
  }
  
  Logger.log("✅ All required sheets loaded successfully.");
  return { sData, eData, dData, statsData };
}
// ✅ CORRECT: This line should only appear ONCE in your file.
const userManualUrl = 'https://docs.google.com/document/d/1BF9XVE1mWOHzXpd9dTHRpcuBkq68FaKmXjt1t_qmyMk/edit?usp=sharing';

/**
 * Creates the 'User Manual' menu in the spreadsheet UI.
 * This should be called from your onOpen() function.
 */
function addUserManualMenu() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('📖 User Manual')
      .addItem('Open User Manual', 'openManualFromMenu') 
      .addToUi();
  } catch (error) {
    Logger.log("Error adding User Manual menu: " + error.message);
  }
}

/**
 * Opens the user manual URL in a new tab.
 * This is the function executed when the menu item is clicked.
 */
function openManualFromMenu() {
  const html = `<script>window.open('${userManualUrl}', '_blank'); google.script.host.close();</script>`;
  const htmlOutput = HtmlService.createHtmlOutput(html);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Opening Manual...');
}

