/* global SpreadsheetApp, g */
/* exported fShowMessage, fLoadSheetToArray, fNormalizeTags */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - User Interface Utilities
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fShowMessage
   Purpose: Displays a simple modal pop-up message to the user.
   Assumptions: None.
   Notes: This is our standard method for all user-facing modal alerts.
   @param {string} title - The title to display in the message box header.
   @param {string} message - The main body of the message.
   @returns {void}
*/
function fShowMessage(title, message) {
  SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
} // End function fShowMessage

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - User Interface Utilities
// Start - Data Handling Utilities
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fParseA1Notation
   Purpose: Parses a custom A1 notation string into an object of rows and columns.
   Assumptions: The input string format is "A,1,3-4,D-F".
   Notes: This is the core parser for the Show/Hide All feature.
   @param {string} notationString - The string to parse.
   @returns {{rows: number[], cols: number[]}} An object containing arrays of row and column numbers.
*/
function fParseA1Notation(notationString) {
  const output = { rows: [], cols: [] };
  if (!notationString) return output;

  const parts = notationString.split(',');

  parts.forEach(part => {
    // Handle row ranges (e.g., "3-5")
    if (part.includes('-') && !isNaN(part.split('-')[0])) {
      const [start, end] = part.split('-').map(Number);
      for (let i = start; i <= end; i++) {
        output.rows.push(i);
      }
    }
    // Handle single rows
    else if (!isNaN(part)) {
      output.rows.push(Number(part));
    }
    // Handle column ranges (e.g., "D-F")
    else if (part.includes('-')) {
      const [start, end] = part.split('-').map(p => p.toUpperCase().charCodeAt(0));
      for (let i = start; i <= end; i++) {
        output.cols.push(i - 64);
      }
    }
    // Handle single columns
    else {
      output.cols.push(part.toUpperCase().charCodeAt(0) - 64);
    }
  });

  // Remove duplicates and sort
  output.rows = [...new Set(output.rows)].sort((a, b) => a - b);
  output.cols = [...new Set(output.cols)].sort((a, b) => a - b);

  return output;
} // End function fParseA1Notation



/* function fLoadSheetToArray
   Purpose: Loads an entire sheet's data into the global g object for in-memory processing.
   Assumptions: The sheet with the specified sheetName exists in the provided spreadsheet object.
   Notes: Creates the necessary object structure within g if it doesn't exist.
   @param {string} spreadsheetName - The key to use for the spreadsheet in the g object (e.g., 'Ver').
   @param {string} sheetName - The exact, case-sensitive name of the sheet to load.
   @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss=SpreadsheetApp.getActiveSpreadsheet()] - The spreadsheet object to load from.
   @returns {void}
*/
function fLoadSheetToArray(spreadsheetName, sheetName, ss = SpreadsheetApp.getActiveSpreadsheet()) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" could not be found.`);
  }

  // Ensure the object structure exists
  if (!g[spreadsheetName]) {
    g[spreadsheetName] = {};
  }
  if (!g[spreadsheetName][sheetName]) {
    g[spreadsheetName][sheetName] = {};
  }

  g[spreadsheetName][sheetName].arr = sheet.getDataRange().getValues();
} // End function fLoadSheetToArray

/* function fNormalizeTags
   Purpose: Converts a raw tag string into an array of standardized tags.
   Assumptions: None.
   Notes: Processes tags to be case-insensitive, space-insensitive, and comma-separated.
   @param {string} tagString - The raw string from a tag cell (e.g., "Character Name, ID").
   @returns {string[]} An array of normalized tags (e.g., ['charactername', 'id']).
*/
function fNormalizeTags(tagString) {
  if (!tagString || typeof tagString !== 'string') {
    return [];
  }
  return tagString
    .toLowerCase()
    .replace(/\s+/g, '')
    .split(',')
    .filter(tag => tag); // Filter out any empty strings that result from ",," or trailing commas
} // End function fNormalizeTags