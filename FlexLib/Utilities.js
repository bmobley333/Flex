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

/* function fLoadSheetToArray
   Purpose: Loads an entire sheet's data into the global g object for in-memory processing.
   Assumptions: The active spreadsheet contains a sheet with the specified sheetName.
   Notes: Creates the necessary object structure within g if it doesn't exist.
   @param {string} spreadsheetName - The key to use for the spreadsheet in the g object (e.g., 'Ver').
   @param {string} sheetName - The exact, case-sensitive name of the sheet to load.
   @returns {void}
*/
function fLoadSheetToArray(spreadsheetName, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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