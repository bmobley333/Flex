/* global g, PropertiesService, SpreadsheetApp, fBuildTagMaps, fLoadSheetToArray */
/* exported fGetSheetId, fLoadSheetIDsFromCodex */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - ID Management & Caching
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fGetSheetId
   Purpose: Gets a specific spreadsheet ID, using a cache-first approach for performance.
   Assumptions: The Codex sheet is set up with a 'Versions' sheet.
   Notes: This is the primary function for retrieving any versioned sheet ID.
   @param {string} version - The version number of the sheet ID to retrieve (e.g., '3').
   @param {string} ssAbbr - The abbreviated name of the sheet (e.g., 'CS').
   @returns {string} The spreadsheet ID.
*/
function fGetSheetId(version, ssAbbr) {
  // 1. Check if the in-memory cache is populated.
  if (Object.keys(g.sheetIDs).length === 0) {
    fLoadSheetIDsFromStorage();
  }

  // 2. If the cache is still empty, load from the Codex sheet as a last resort.
  if (Object.keys(g.sheetIDs).length === 0) {
    fLoadSheetIDsFromCodex();
    fCacheSheetIDsToStorage();
  }

  if (g.sheetIDs[version] && g.sheetIDs[version][ssAbbr]) {
    return g.sheetIDs[version][ssAbbr].ssid;
  } else {
    throw new Error(`Could not find Sheet ID for version "${version}", abbreviation "${ssAbbr}".`);
  }
} // End function fGetSheetId

/* function fLoadSheetIDsFromCodex
   Purpose: Reads the Codex's internal "Versions" sheet to build the initial sheet ID cache.
   Assumptions: The 'Versions' sheet is tagged with 'tablestart', 'tableend', 'version', 'ssabbr', and 'ssid'.
   Notes: This should only ever run once per user.
   @returns {void}
*/
function fLoadSheetIDsFromCodex() {
  const ssKey = 'Codex';
  const sheetName = 'Versions';

  fBuildTagMaps(ssKey, sheetName);

  const { arr, rowTags, colTags } = g[ssKey][sheetName];
  const startRow = rowTags.tablestart;
  const endRow = rowTags.tableend;

  if (typeof startRow === 'undefined' || typeof endRow === 'undefined') {
    throw new Error(`Could not find 'tablestart' or 'tableend' row tags in the "${sheetName}" sheet.`);
  }

  for (let r = startRow; r <= endRow; r++) {
    const version = String(arr[r][colTags.version]); // This line is changed
    const abbr = arr[r][colTags.ssabbr];
    const id = arr[r][colTags.ssid];
    const fullName = arr[r][colTags.ssfullname];

    if (!version || !abbr || !id) continue; // Skip incomplete rows

    if (!g.sheetIDs[version]) {
      g.sheetIDs[version] = {};
    }

    g.sheetIDs[version][abbr] = {
      version: version,
      ssabbr: abbr,
      ssid: id,
      ssfullname: fullName,
    };
  }
} // End function fLoadSheetIDsFromCodex

/* function fLoadSheetIDsFromStorage
   Purpose: Loads the sheet ID cache from persistent PropertiesService into the global g object.
   Assumptions: None.
   Notes: Internal helper function.
   @returns {void}
*/
function fLoadSheetIDsFromStorage() {
  const properties = PropertiesService.getScriptProperties();
  const storedIDs = properties.getProperty('sheetIDs');
  if (storedIDs) {
    g.sheetIDs = JSON.parse(storedIDs);
  }
} // End function fLoadSheetIDsFromStorage

/* function fCacheSheetIDsToStorage
   Purpose: Saves the in-memory sheet ID cache to persistent PropertiesService.
   Assumptions: g.sheetIDs has been populated.
   Notes: Internal helper function.
   @returns {void}
*/
function fCacheSheetIDsToStorage() {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty('sheetIDs', JSON.stringify(g.sheetIDs));
} // End function fCacheSheetIDsToStorage