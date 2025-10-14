/* global g, fLoadSheetToArray, fBuildTagMaps, SpreadsheetApp */
/* exported fGetSheetData */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Data Caching & Retrieval
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


/* function fInvalidateSheetCache
   Purpose: Invalidates the session cache for a specific sheet.
   Assumptions: None.
   Notes: Called by triggers when a sheet's structure changes, forcing fGetSheetData to re-read it on the next call.
   @param {string} ssKey - The key for the spreadsheet in the g object (e.g., 'CS').
   @param {string} sheetName - The exact, case-sensitive name of the sheet to invalidate.
   @returns {void}
*/
function fInvalidateSheetCache(ssKey, sheetName) {
  if (g[ssKey] && g[ssKey][sheetName]) {
    delete g[ssKey][sheetName];
  }
} // End function fInvalidateSheetCache

/* function fGetSheetData
   Purpose: The master gatekeeper for retrieving sheet data, using a lazy-loading session cache.
   Assumptions: None.
   Notes: This is the central function for all sheet data access, ensuring performance and data integrity.
   @param {string} ssKey - The key for the spreadsheet in the g object (e.g., 'Codex', 'CS').
   @param {string} sheetName - The exact, case-sensitive name of the sheet to load.
   @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss=SpreadsheetApp.getActiveSpreadsheet()] - The spreadsheet object to load from. Defaults to the active one.
   @param {boolean} [forceRefresh=false] - If true, ignores the cache and re-reads the data from the sheet.
   @returns {object} The sheet data object, containing { arr, rowTags, colTags }.
*/
function fGetSheetData(ssKey, sheetName, ss = SpreadsheetApp.getActiveSpreadsheet(), forceRefresh = false) {
  // 1. Check if the cache exists and a refresh is NOT forced.
  if (!forceRefresh && g[ssKey] && g[ssKey][sheetName] && g[ssKey][sheetName].arr) {
    return g[ssKey][sheetName]; // Return the cached data instantly.
  }

  // 2. If cache is empty or a refresh is forced, load the data from the spreadsheet.
  fLoadSheetToArray(ssKey, sheetName, ss);
  fBuildTagMaps(ssKey, sheetName);

  // 3. Return the newly loaded data.
  return g[ssKey][sheetName];
} // End function fGetSheetData