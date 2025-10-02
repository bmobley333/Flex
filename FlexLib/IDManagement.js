/* global g, PropertiesService, SpreadsheetApp, fBuildTagMaps, fLoadSheetToArray */
/* exported fGetSheetId, fLoadSheetIDsFromMyVersions */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - ID Management & Caching
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fGetSheetId
   Purpose: Gets a specific spreadsheet ID from the player's local collection, using a cache-first approach.
   Assumptions: The Codex has a <MyVersions> sheet that has been populated.
   Notes: This is the primary function for retrieving any local versioned sheet ID.
   @param {string} version - The version number of the sheet ID to retrieve (e.g., '3').
   @param {string} ssAbbr - The abbreviated name of the sheet (e.g., 'CS').
   @returns {string} The spreadsheet ID.
*/
function fGetSheetId(version, ssAbbr) {
  // 1. Check if the in-memory cache is populated.
  if (Object.keys(g.sheetIDs).length === 0) {
    fLoadSheetIDsFromStorage();
  }

  // 2. If the cache is still empty, load from the <MyVersions> sheet as a last resort.
  if (Object.keys(g.sheetIDs).length === 0) {
    fLoadSheetIDsFromMyVersions();
    fCacheSheetIDsToStorage();
  }

  if (g.sheetIDs[version] && g.sheetIDs[version][ssAbbr]) {
    return g.sheetIDs[version][ssAbbr].ssid;
  } else {
    // If not found, force a reload from the sheet and try one more time.
    fLoadSheetIDsFromMyVersions();
    fCacheSheetIDsToStorage();
    if (g.sheetIDs[version] && g.sheetIDs[version][ssAbbr]) {
      return g.sheetIDs[version][ssAbbr].ssid;
    } else {
      throw new Error(`Could not find a local Sheet ID for version "${version}", abbreviation "${ssAbbr}".`);
    }
  }
} // End function fGetSheetId

/* function fLoadSheetIDsFromMyVersions
   Purpose: Reads the Codex's <MyVersions> sheet to build the cache of local file IDs.
   Assumptions: The <MyVersions> sheet is tagged with 'tablestart', 'tableend', 'version', 'ssabbr', and 'ssid'.
   Notes: This powers the cache with the player's own file data.
   @returns {void}
*/
function fLoadSheetIDsFromMyVersions() {
  const ssKey = 'Codex';
  const sheetName = 'MyVersions';

  fBuildTagMaps(ssKey, sheetName);

  const { arr, rowTags, colTags } = g[ssKey][sheetName];
  const startRow = rowTags.tablestart;
  const endRow = rowTags.tableend;

  if (typeof startRow === 'undefined' || typeof endRow === 'undefined') {
    throw new Error(`Could not find 'tablestart' or 'tableend' row tags in the "${sheetName}" sheet.`);
  }

  // Clear the in-memory cache before reloading
  g.sheetIDs = {};

  for (let r = startRow; r <= endRow; r++) {
    const version = String(arr[r][colTags.version]);
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
} // End function fLoadSheetIDsFromMyVersions

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