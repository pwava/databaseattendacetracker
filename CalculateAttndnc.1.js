/**
 * Extracts a purely numeric ID if the input is already a number or a string representing a non-negative integer.
 * It NO LONGER parses "BEL" prefixes. Any string containing non-numeric characters (like "BEL123")
 * will be treated as an invalid/missing ID.
 * @param {any} codeValue The value from the sheet that might be a numeric ID.
 * @returns {number|null} The extracted number, or null if not a valid purely numeric ID.
 */
function extractNumericBel(codeValue) {
  if (typeof codeValue === 'number' && Number.isInteger(codeValue) && codeValue >= 0) {
    return codeValue;
  }
  if (typeof codeValue === 'string') {
    const numStr = codeValue.trim();

    if (numStr === '') {
      return null;
    }

    if (!/^\d+$/.test(numStr)) {
      Logger.log(`ℹ️ extractNumericBel: Input "${codeValue}" is not a purely numeric string. Will treat as missing/invalid ID.`);
      return null;
    }

    const num = parseInt(numStr, 10);
    if (!isNaN(num) && num >= 0) {
      return num;
    } else {
      Logger.log(`⚠️ extractNumericBel: Could not parse valid non-negative number from numeric string "${numStr}". Parsed: ${num}`);
      return null;
    }
  }
  return null;
}

/**
 * Fetches data using getDataFromSheets, matches/assigns purely NUMERIC IDs,
 * and formats attendance records into a standardized 11-column structure.
 *
 * This function first collects all existing PURELY NUMERIC IDs from the Directory,
 * Event Attendance (Column A), and Service Attendance (Column A) sheets using the modified extractNumericBel.
 * It builds a comprehensive name-to-NUMERIC_ID mapping (belMap). It then finds
 * the highest number among these numeric codes and generates new unique NUMERIC IDs
 * sequentially from the next number when a person doesn't have an existing valid numeric code.
 *
 * Assumes data object from getDataFromSheets contains:
 * - sData: Array of rows from Service Attendance
 * - eData: Array of rows from Event Attendance
 * - dData: Array of rows from Directory (now correctly assumed ID in row[0], Name in row[1])
 *
 * The function returns an array of arrays, where each inner array has
 * at least 11 elements, conforming to the "Event Attendance" column structure:
 * [0: Numeric ID (number), 1: Full Name, ..., 10: Timestamp]
 *
 * @returns {Array<Array<any>>} An array of formatted attendance records, or empty array if data loading fails.
 */
function matchOrAssignBelCodes() {
  const data = getDataFromSheets();
  if (!data) {
    Logger.log("❌ No data loaded from sheets in matchOrAssignBelCodes. Exiting.");
    return [];
  }

  const { sData, eData, dData } = data;

  const belMap = new Map();
  const allUsedCodes = new Set();
  const normalize = name => name?.toString().trim().toLowerCase();

  // --- Step 1: Populate belMap and allUsedCodes (with NUMBERS) from the Directory sheet ---
  // FIX APPLIED HERE: Corrected column indices for Directory data (dData)
  if (dData && dData.length > 1) {
    dData.slice(1).forEach((row, index) => {
      // Based on your confirmation:
      // Column A (index 0): Person ID
      // Column B (index 1): Full Name
      if (row.length > 1) { // Ensure at least Column A (ID) and Column B (Full Name) exist
        const originalBel = row[0]; // **FIXED: Person ID from Column A (index 0)**
        const name = normalize(row[1]); // **FIXED: Full Name from Column B (index 1)**
        const numericBel = extractNumericBel(originalBel);

        if (name) {
          if (numericBel !== null) {
            if (!belMap.has(name)) {
              belMap.set(name, numericBel);
            }
            allUsedCodes.add(numericBel);
          } else if (originalBel && originalBel.toString().trim() !== '') {
            // This log should now ONLY appear if column A of your Directory has non-numeric (non-ID) data
            Logger.log(`ℹ️ Directory: Row ${index + 2}: Value "${originalBel}" in Person ID column is not a plain number and will be ignored. A new numeric ID may be generated for "${name}" if needed.`);
          }
        } else {
          Logger.log(`⚠️ Directory: Row ${index + 2}: Skipping row due to missing Full Name in Column B. Row data: ${JSON.stringify(row)}`);
        }
      } else {
        Logger.log(`⚠️ Directory: Row ${index + 2}: Skipping row due to insufficient columns (${row.length} found). Expected at least 2 (Person ID, Full Name). Row data (partial): ${JSON.stringify(row.slice(0,2))}`);
      }
    });
    Logger.log(`✅ Populated numeric ID map and used codes from Directory (${dData.length > 1 ? dData.length - 1 : 0} data rows processed).`);
  } else {
    Logger.log("⚠️ matchOrAssignBelCodes: Directory data (dData) is empty or missing headers.");
  }

  // --- Step 2: Add existing NUMERIC IDs from Attendance sheets and update belMap ---
  const attendanceDataRaw = [];
  if (eData && eData.length > 1) attendanceDataRaw.push(...eData.slice(1));
  if (sData && sData.length > 1) attendanceDataRaw.push(...sData.slice(1));

  attendanceDataRaw.forEach((row, idx) => {
    if (row.length > 1) {
      // Column A (Person ID) and Column B (Full Name) in attendance sheets are correctly read here
      const originalBelFromRow = row[0];
      const name = normalize(row[1]);
      const numericBelFromRow = extractNumericBel(originalBelFromRow);

      if (numericBelFromRow !== null) {
        allUsedCodes.add(numericBelFromRow);
        if (name && !belMap.has(name)) {
          belMap.set(name, numericBelFromRow);
        }
      } else if (originalBelFromRow && originalBelFromRow.toString().trim() !== '' && name) {
          Logger.log(`ℹ️ Attendance Sheets (Source Row approx. ${idx + 1}): Value "${originalBelFromRow}" for name "${name}" in ID column is not a plain number and will be ignored. A new ID may be generated.`);
      }
    }
  });
  Logger.log(`✅ Added existing numeric IDs from Attendance sheets to used codes set. Total unique used NUMERIC codes found: ${allUsedCodes.size}. Total NUMERIC ID mappings found: ${belMap.size}`);

  // --- Step 3: Initialize NUMERIC ID Code generator ---
  let highestFoundNum = 0;
  allUsedCodes.forEach(code => {
    if (typeof code === 'number' && code > highestFoundNum) {
      highestFoundNum = code;
    }
  });

  let codeCounter = highestFoundNum + 1;
  Logger.log(`✅ Initialized NUMERIC ID code counter to: ${codeCounter}. Highest valid number found among existing ID codes was ${highestFoundNum}.`);

  // --- Step 4: Define the function to generate the next unique NUMERIC ID code ---
  const generateBEL = () => {
    while (true) {
      const currentNumericCode = codeCounter;
      if (!allUsedCodes.has(currentNumericCode)) {
        allUsedCodes.add(currentNumericCode);
        codeCounter++;
        return currentNumericCode;
      }
      codeCounter++;
      if (codeCounter > highestFoundNum + 20000) {
        Logger.log(`ERROR: ID code counter (${codeCounter}) significantly exceeds highest found number (${highestFoundNum}). Potential issue.`);
        throw new Error("ID code counter exceeded a safe threshold, potential infinite loop or too many users without pre-existing IDs.");
      }
    }
  };

  // --- Step 5: Process Attendance Data and Assign Final NUMERIC IDs ---
  const results = [];
  attendanceDataRaw.forEach(row => {
    if (row.length < 2) {
      Logger.log(`⚠️ Processing Attendance: Skipping row due to insufficient columns for Name. Row data: ${JSON.stringify(row)}`);
      return;
    }
    const name = normalize(row[1]);
    if (!name) {
      Logger.log(`⚠️ Processing Attendance: Skipping row due to missing Name in Column B. Row data: ${JSON.stringify(row)}`);
      return;
    }

    let numericBel;
    if (belMap.has(name)) {
      numericBel = belMap.get(name);
    } else {
      numericBel = generateBEL();
      belMap.set(name, numericBel);
      Logger.log(`✅ Generated new NUMERIC ID ${numericBel} for name "${row[1]}".`);
    }

    let formattedRow = Array(11).fill("");
    formattedRow[0] = numericBel;

    if (row.length >= 11 && typeof row[10] !== 'undefined') {
      formattedRow[1] = row[1];
      formattedRow[2] = row[2];
      formattedRow[3] = row[3];
      formattedRow[4] = row[4];
      formattedRow[5] = row[5];
      formattedRow[6] = row[6];
      formattedRow[7] = row[7];
      formattedRow[8] = row[8];
      formattedRow[9] = row[9];
      formattedRow[10] = row[10];
    } else if (row.length >= 8 && typeof row[4] !== 'undefined') {
      formattedRow[1] = row[1];
      formattedRow[2] = "Sunday Service";
      formattedRow[3] = "Service";
      formattedRow[4] = row[2];
      formattedRow[5] = row[3];
      formattedRow[6] = row[6];
      formattedRow[9] = "";
      formattedRow[10] = row[4];
    } else {
      Logger.log(`⚠️ Processing Attendance: Skipping row for "${name}" (ID: ${numericBel}) with unrecognized structure. Row data: ${JSON.stringify(row)}`);
      return;
    }
    results.push(formattedRow);
  });

  Logger.log(`✅ Total attendance records matched and formatted with NUMERIC IDs: ${results.length}`);
  return results;
}

/**
 * Helper function to retrieve all data from the "Service Attendance" sheet.
 * Assumes the sheet name is exactly "Service Attendance".
 *
 * @returns {Array<Array<any>>} An array of arrays representing the data, or empty array if sheet not found.
 */
function getServiceAttendanceData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "Service Attendance"; // Confirmed: This is the name of your Service Attendance tab
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log(`❌ Sheet "${sheetName}" not found. Please ensure the tab name is exact.`);
    return [];
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow <= 1) {
    Logger.log(`⚠️ No data found in "${sheetName}" sheet (only header row or empty data range).`);
    return [];
  }

  const dataRange = sheet.getRange(2, 1, lastRow - 1, lastColumn);
  const data = dataRange.getValues();

  Logger.log(`✅ Retrieved ${data.length} rows from "${sheetName}" sheet.`);
  return data;
}

/**
 * Helper function to retrieve all data from the specified sheets.
 * It now distinguishes between the active spreadsheet for attendance data
 * and an external spreadsheet for directory data, whose ID is fetched from PropertiesService.
 *
 * @returns {Object} An object containing arrays of data for dData, eData, and sData.
 */
function getDataFromSheets() {
  Logger.log("ℹ️ getDataFromSheets: Fetching data from spreadsheet sheets, including external directory.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const scriptProperties = PropertiesService.getScriptProperties();
  const directorySheetId = scriptProperties.getProperty("DIRECTORY_SPREADSHEET_ID"); // <<< CORRECT KEY NAME

  let externalDirectorySs = null;
  if (directorySheetId) {
    try {
      externalDirectorySs = SpreadsheetApp.openById(directorySheetId);
      Logger.log(`✅ Opened external Directory spreadsheet with ID: ${directorySheetId}`);
    } catch (e) {
      Logger.log(`❌ Failed to open external Directory spreadsheet with ID "${directorySheetId}": ${e.message}`);
      // Proceed without external directory data if it fails
    }
  } else {
    Logger.log("⚠️ No 'DIRECTORY_SPREADSHEET_ID' found in script properties. Directory data will be empty.");
  }

  const getSheetData = (sheetName, spreadsheet = ss) => {
    const targetSs = spreadsheet;
    if (!targetSs) {
      Logger.log(`❌ Cannot get data from sheet "${sheetName}" because the target spreadsheet is not available.`);
      return [];
    }

    const sheet = targetSs.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`❌ Sheet "${sheetName}" not found in spreadsheet "${targetSs.getName()}". Returning empty array.`);
      return [];
    }
    const data = sheet.getDataRange().getValues();
    Logger.log(`✅ Retrieved ${data.length} rows from "${sheetName}" in "${targetSs.getName()}".`);
    return data;
  };

  return {
    dData: externalDirectorySs ? getSheetData("Sunday Registration", externalDirectorySs) : [],
    eData: getSheetData("Event Attendance", ss),
    sData: getSheetData("Service Attendance", ss)
  };
}

// --- NEW HELPER: Retrieves current Person ID and ALL relevant stats from Attendance Stats sheet for comparison ---
/**
 * Reads the current data from the 'Attendance Stats' sheet (Columns A, E, F, G, H, I, J, K)
 * to get existing statistics values for comparison.
 * @returns {Map<number, Object>} A map where key is Person ID (numeric) and value is an object
 * containing existing stats (quarter, month, volunteer, lastDate, lastEvent, total, lastYearServiceCount).
 */
function getExistingAttendanceStatsCounts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance Stats");
  const existingStats = new Map(); // Map: Person ID -> {quarter, month, volunteer, lastDate, lastEvent, total, lastYearServiceCount}

  if (!sheet || sheet.getLastRow() <= 1) {
    Logger.log("⚠️ 'Attendance Stats' sheet is empty or not found for checking existing stats.");
    return existingStats;
  }

  // Get data from Column A (index 0) to Column K (index 10)
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11);
  const values = dataRange.getValues();

  values.forEach(row => {
    const personId = extractNumericBel(row[0]); // Column A: Person ID

    if (personId !== null) {
      existingStats.set(personId, {
        quarter: typeof row[4] === 'number' ? row[4] : 0,   // Column E: Quarter Events
        month: typeof row[5] === 'number' ? row[5] : 0,     // Column F: Month Events
        volunteer: typeof row[6] === 'number' ? row[6] : 0, // Column G: Volunteer Count
        lastDate: row[7] instanceof Date ? row[7] : null,   // Column H: Last Date (as Date object)
        lastEvent: String(row[8] || ''),                   // Column I: Last Event Name
        total: typeof row[9] === 'number' ? row[9] : 0,     // Column J: Total Unique Events
        lastYearServiceCount: typeof row[10] === 'number' ? row[10] : 0 // Column K: Last Year Service Count
      });
    }
  });
  Logger.log(`✅ Retrieved ${existingStats.size} existing attendance stats records from 'Attendance Stats' sheet for comparison.`);
  return existingStats;
}
// --- END NEW HELPER ---


/**
 * Calculates attendance statistics based on formatted raw data.
 * Groups entries by BEL code and summarizes attendance
 * for the current month, quarter, and year,
 * including volunteer instances and last attended date.
 *
 * @returns {Array<Array<any>>} An array of arrays containing summarized attendance statistics per individual, or empty array if no data to process.
 */
function calculateAttendanceStats() {
  const rawData = matchOrAssignBelCodes();

  if (!rawData || rawData.length === 0) {
    Logger.log("❌ No data received from matchOrAssignBelCodes or data is empty after formatting.");
    return [];
  }

  const now = new Date();
  const currentMonth = now.getMonth();
  const currentQuarter = Math.floor(currentMonth / 3);
  const currentYear = now.getFullYear();

  const previousCalendarYear = currentYear - 1;
  Logger.log(`Debug: Calculating last year's attendance for calendar year: ${previousCalendarYear}`);

  const serviceAttendanceData = getServiceAttendanceData();
  const serviceAttendanceLastYear = new Map(); // Map: Person ID (numeric) -> Set of unique date strings (YYYY-MM-DD)

  if (serviceAttendanceData && serviceAttendanceData.length > 0) {
    serviceAttendanceData.forEach(row => {
      const personId = extractNumericBel(row[0]);
      const dateStr = row[4];

      let serviceDate;
      if (dateStr instanceof Date) {
        serviceDate = dateStr;
      } else if (typeof dateStr === 'number') {
        serviceDate = new Date((dateStr - (25567 + 2)) * 86400 * 1000);
      } else {
        serviceDate = new Date(String(dateStr));
      }

      if (isNaN(serviceDate.getTime())) {
        Logger.log(`⚠️ Service Attendance: Skipping invalid date "${dateStr}" for Person ID ${row[0]}.`);
        return;
      }

      if (personId !== null && serviceDate.getFullYear() === previousCalendarYear) {
        if (!serviceAttendanceLastYear.has(personId)) {
          serviceAttendanceLastYear.set(personId, new Set());
        }
        const uniqueDateKey = serviceDate.toISOString().split('T')[0];
        serviceAttendanceLastYear.get(personId).add(uniqueDateKey);
      }
    });
  }
  Logger.log(`✅ Processed Service Attendance for previous calendar year (${previousCalendarYear}). Found data for ${serviceAttendanceLastYear.size} individuals.`);

  const existingStatsFromSheet = getExistingAttendanceStatsCounts(); // Renamed for clarity

  const grouped = new Map();

  rawData.forEach(row => {
    if (row.length < 11) {
      Logger.log(`⚠️ calculateAttendanceStats: Skipping row due to insufficient columns (${row.length} found). Expected at least 11. Row data (partial): ${JSON.stringify(row.slice(0, 11))}`);
      return;
    }

    const bel = row[0];
    const name = row[1];
    const eventName = row[2];
    const eventId = row[3];
    const role = row[9];
    const dateStr = row[10];

    let date;
    if (dateStr instanceof Date) {
      date = dateStr;
    } else if (typeof dateStr === 'number') {
      date = new Date((dateStr - (25567 + 2)) * 86400 * 1000);
    } else {
      date = new Date(String(dateStr));
    }

    if (isNaN(date.getTime())) {
      Logger.log(`⚠️ Skipping invalid date: "${dateStr}" found for BEL ${bel}. Full row data: ${JSON.stringify(row)}`);
      return;
    }

    const isSundayService = typeof eventName === 'string' && /sunday service/i.test(eventName);
    const isVolunteer = typeof role === 'string' && String(role).toLowerCase().includes("volunteer");
    const eventNameKey = typeof eventName === 'string' ? eventName : 'UnknownEvent';
    const eventIdKey = typeof eventId === 'string' ? eventId : 'UnknownID';
    const eventKey = isSundayService ? `sunday service-${date.toDateString()}` : `${eventNameKey}-${eventIdKey}`;

    const record = {
      name,
      date,
      eventKey,
      month: date.getMonth(),
      quarter: Math.floor(date.getMonth() / 3),
      year: date.getFullYear(),
      isVolunteer,
      isSundayService,
    };

    const belString = String(bel);
    if (!grouped.has(belString)) {
      grouped.set(belString, []);
    }
    grouped.get(belString).push(record);
  });

  const summary = [];

  grouped.forEach((records, bel) => {
    const uniqueEvents = new Set();
    const monthEvents = new Set();
    const quarterEvents = new Set();
    let volunteerCount = 0;

    records.forEach(r => {
      uniqueEvents.add(r.eventKey);
      if (r.year === currentYear && r.month === currentMonth) {
        monthEvents.add(r.eventKey);
      }
      if (r.year === currentYear && r.quarter === currentQuarter) {
        quarterEvents.add(r.eventKey);
      }
      if (r.isVolunteer && r.year === currentYear) {
        volunteerCount++;
      }
    });

    records.sort((a, b) => b.date.getTime() - a.date.getTime());
    const mostRecentRecord = records.length > 0 ? records[0] : null;

    let fullName = '';
    let lastDate = '';
    let lastEventName = '';

    if (mostRecentRecord) {
      fullName = mostRecentRecord.name;
      lastDate = mostRecentRecord.date;
      const lastEventKey = mostRecentRecord.eventKey;
      const lastEventParts = lastEventKey.split('-');
      lastEventName = lastEventParts.length > 0 ? lastEventParts[0] : lastEventKey;
    }

    const totalUniqueEvents = uniqueEvents.size;

    const belAsNumber = extractNumericBel(bel);
    let lastYearServiceCount = 0;

    if (belAsNumber !== null && serviceAttendanceLastYear.has(belAsNumber)) {
      lastYearServiceCount = serviceAttendanceLastYear.get(belAsNumber).size;
    } else if (belAsNumber === null) {
      Logger.log(`⚠️ Person ID (BEL from match script) "${bel}" is not a valid numeric ID. Last Year Service Attendance will be 0 for "${fullName}".`);
    } else {
       Logger.log(`ℹ️ No last year service attendance found for Person ID ${belAsNumber} ("${fullName}") for calendar year ${previousCalendarYear}. Count will be 0.`);
    }

    // --- NEW ADDITION: Define current calculated stats variables for logging and summary push ---
    const _currentQuarterValue = quarterEvents.size;
    const _currentMonthValue = monthEvents.size;
    const _currentVolunteerValue = volunteerCount;
    const _currentTotalValue = totalUniqueEvents;
    const _currentLastYearServiceCountValue = lastYearServiceCount;
    const _currentLastDateFormattedValue = lastDate instanceof Date ? Utilities.formatDate(lastDate, Session.getScriptTimeZone(), "MM/dd/yyyy") : String(lastDate || '');
    const _currentLastEventValue = lastEventName; // This is the raw string
    // --- END NEW ADDITION ---

    // --- NEW ADDITION: Detailed Logging for ALL updates ---
    if (belAsNumber !== null) {
      const oldStats = existingStatsFromSheet.get(belAsNumber);

      if (!oldStats) {
        // New record
        Logger.log(`✨ NEW RECORD: Person ID ${belAsNumber} ("${fullName}") has been added.`);
        Logger.log(`   Counts: Q:${_currentQuarterValue} M:${_currentMonthValue} V:${_currentVolunteerValue} Total:${_currentTotalValue} LastYear:${_currentLastYearServiceCountValue}`);
        Logger.log(`   Last Attended: ${_currentLastDateFormattedValue} (${_currentLastEventValue})`);

      } else {
        let changes = [];
        if (oldStats.quarter !== _currentQuarterValue) changes.push(`Q:${oldStats.quarter} -> ${_currentQuarterValue}`);
        if (oldStats.month !== _currentMonthValue) changes.push(`M:${oldStats.month} -> ${_currentMonthValue}`);
        if (oldStats.volunteer !== _currentVolunteerValue) changes.push(`V:${oldStats.volunteer} -> ${_currentVolunteerValue}`);
        if (oldStats.total !== _currentTotalValue) changes.push(`Total:${oldStats.total} -> ${_currentTotalValue}`);
        if (oldStats.lastYearServiceCount !== _currentLastYearServiceCountValue) changes.push(`LastYear:${oldStats.lastYearServiceCount} -> ${_currentLastYearServiceCountValue}`);

        // Compare Date objects using getTime() or formatted string for consistency
        const oldLastDateFormatted = oldStats.lastDate instanceof Date ? Utilities.formatDate(oldStats.lastDate, Session.getScriptTimeZone(), "MM/dd/yyyy") : String(oldStats.lastDate || '');
        if (oldLastDateFormatted !== _currentLastDateFormattedValue) changes.push(`LastDate:'${oldLastDateFormatted}' -> '${_currentLastDateFormattedValue}'`);
        if (oldStats.lastEvent !== _currentLastEventValue) changes.push(`LastEvent:'${oldStats.lastEvent}' -> '${_currentLastEventValue}'`);

        if (changes.length > 0) {
          Logger.log(`🔄 UPDATED RECORD: Person ID ${belAsNumber} ("${fullName}") had changes:`);
          changes.forEach(change => Logger.log(`   - ${change}`));
        } else {
          // If no changes, no specific log unless very verbose debugging is needed
          // Logger.log(`Debug: Person ID ${belAsNumber} ("${fullName}") stats are unchanged.`);
        }
      }
    }
    // --- END NEW ADDITION ---

    Logger.log(`Debug: Person ID (from match script): ${bel}, Full Name: ${fullName}, Last Year Service Attendance Count: ${_currentLastYearServiceCountValue}`);

    summary.push([
      bel,
      fullName,
      "",
      "",
      _currentQuarterValue,
      _currentMonthValue,
      _currentVolunteerValue,
      lastDate, // Pass the raw Date object
      lastEventName,
      _currentTotalValue,
      _currentLastYearServiceCountValue
    ]);
  });

  // --- MODIFIED: Sort the summary array NUMERICALLY by Person ID (index 0) ---
  summary.sort((a, b) => {
    const idA = a[0]; // Person ID is at index 0
    const idB = b[0]; // Person ID is at index 0
    return idA - idB;
  });
  Logger.log("✅ Summary data sorted numerically by Person ID.");

  Logger.log("✅ Attendance stats calculated for: " + summary.length + " individuals.");

  return summary;
}

function updateAttendanceStatsSheet() {
  const finalData = calculateAttendanceStats();

  if (!finalData || finalData.length === 0) {
    Logger.log("❌ No final data to update the 'Attendance Stats' sheet.");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance Stats");
  if (!sheet) {
    Logger.log("❌ 'Attendance Stats' sheet not found. Cannot update.");
    return;
  }

  // Format the data for writing to the sheet
  const output = finalData.map(row => {
    const [
      bel,
      fullName,
      , // Placeholder for First Name (will be derived)
      , // Placeholder for Last Name (will be derived)
      quarter,
      month,
      volunteer,
      lastDate,
      lastEvent,
      total,
      lastYearServiceCount // Grab the 11th element (Column K data)
    ] = row;

    const nameParts = fullName ? String(fullName).trim().split(/\s+/) : [];
    const firstName = nameParts.length > 0 ? nameParts[0] : "";
    const lastName = nameParts.length > 1 ? nameParts.slice(1).join(" ") : "";

    let formattedDate = "";
    if (lastDate instanceof Date && !isNaN(lastDate.getTime())) {
      formattedDate = Utilities.formatDate(lastDate, Session.getScriptTimeZone(), "MM/dd/yyyy");
    } else if (lastDate) {
      formattedDate = String(lastDate);
    }

    return [
      bel,
      fullName,
      firstName,
      lastName,
      quarter,
      month,
      volunteer,
      formattedDate,
      lastEvent,
      total,
      lastYearServiceCount // Add the calculated value to the output row
    ];
  });

  const numRows = output.length;
  const numCols = output[0] ? output[0].length : 0;

  // Explicitly clear only up to column K (11 columns)
  const targetNumCols = 11;
  sheet.getRange(2, 1, sheet.getLastRow(), targetNumCols).clearContent();

  // Write the data starting at row 2, column 1
  if (numRows > 0 && numCols > 0) {
    sheet.getRange(2, 1, numRows, numCols).setValues(output);
    Logger.log(`✅ Wrote ${numRows} rows and ${numCols} columns to 'Attendance Stats'.`);
  } else {
    Logger.log(`⚠️ No data to write to 'Attendance Stats' sheet after formatting.`);
  }

}