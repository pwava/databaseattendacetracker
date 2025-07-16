/**
 * @OnlyCurrentDoc
 * This script adds custom menus to the spreadsheet.
 */

// --- Global Constant for the User Manual ---
// By placing this here, it's easy to find and update the URL if it ever changes.
const userManualUrl = 'https://docs.google.com/document/d/1BF9XVE1mWOHzXpd9dTHRpcuBkq68FaKmXjt1t_qmyMk/edit?usp=sharing';

/**
 * Runs when the spreadsheet is opened to add our custom menus.
 * @param {GoogleAppsScript.Events.SheetsOnOpen} e The event object (optional).
 */
function onOpen(e) {
  Logger.log("Master onOpen triggered. AuthMode: " + (e ? e.authMode : 'N/A Event Object'));

  try {
    // --- Main Config Menu ---
    SpreadsheetApp.getUi()
      .createMenu('⚙️ Config')
      .addItem('Set Directory Spreadsheet URL…', 'showDirectoryDialog')
      .addToUi();
    Logger.log("✅ Config menu added by onOpen.");

    // --- User Manual Menu ---
    // This menu uses the one-click method to open a new tab.
    SpreadsheetApp.getUi()
      .createMenu('📖 User Manual')
      .addItem('Open User Manual', 'openManualInNewTab')
      .addToUi();
    Logger.log("✅ User Manual menu added by onOpen.");

  } catch (error) {
    Logger.log("Error adding a menu in onOpen: " + error.message);
  }

  // --- Other Menu Initializations ---
  // These are kept separate as they might have more complex logic.
  try {
    addSundayRegistrationMenu();
    Logger.log("Call to addSundayRegistrationMenu completed from onOpen.");
  } catch (error) {
    Logger.log("Error during addSundayRegistrationMenu in onOpen: " + error.message);
  }

  try {
    addEventRegistrationMenu();
    Logger.log("Call to addEventRegistrationMenu completed from onOpen.");
  } catch (error) {
    Logger.log("Error during addEventRegistrationMenu in onOpen: " + error.message);
  }

  // --- On-Open Data Updates ---
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
 * Opens the user manual URL in a new tab using the one-click workaround.
 * This is the function executed when the menu item is clicked.
 */
function openManualInNewTab() {
  const html = `<script>window.open('${userManualUrl}', '_blank'); google.script.host.close();</script>`;
  const htmlOutput = HtmlService.createHtmlOutput(html).setWidth(100).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Opening...');
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
  const id = extractSpreadsheetId(input);

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
 * • a /d/URL segment, or
 * • a bare ID string
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
 * { sData, eData, dData, statsData },
 * or undefined on failure.
 */
function getDataFromSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
  const eventSheet = ss.getSheetByName("Event Attendance");
  const statsSheet = ss.getSheetByName("Attendance Stats");

  if (!serviceSheet || !eventSheet || !statsSheet) {
    if (!serviceSheet) Logger.log("❌ 'Service Attendance' not found.");
    if (!eventSheet) Logger.log("❌ 'Event Attendance' not found.");
    if (!statsSheet) Logger.log("❌ 'Attendance Stats' not found.");
    return;
  }

  // --- Load data from local sheets ---
  let sData, eData, statsData;
  try {
    sData = serviceSheet.getDataRange().getValues();
    eData = eventSheet.getDataRange().getValues();
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

// NOTE: You will need to have the functions 'addSundayRegistrationMenu', 
// 'addEventRegistrationMenu', and 'updateEventAttendanceCounts' defined elsewhere 
// in your project for the script to run without errors.
