/* global g, PropertiesService, SpreadsheetApp, fBuildTagMaps, fLoadSheetToArray */
/* exported fGetSheetId, fLoadSheetIDsFromMyVersions */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - ID Management & Caching
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fGetMasterSheetId
   Purpose: Gets a specific spreadsheet ID from the master <Versions> sheet.
   Assumptions: The master <Versions> sheet is accessible and correctly tagged.
   Notes: This is used for processes that need direct access to master file IDs, not local copies.
   @param {string} version - The version number of the sheet ID to retrieve (e.g., '3').
   @param {string} ssAbbr - The abbreviated name of the sheet (e.g., 'Tbls').
   @returns {string|null} The spreadsheet ID, or null if not found.
*/
function fGetMasterSheetId(version, ssAbbr) {
  const sourceSS = SpreadsheetApp.openById(g.MASTER_VER_ID);
  fLoadSheetToArray('Ver', 'Versions', sourceSS);
  fBuildTagMaps('Ver', 'Versions');

  const { arr, rowTags, colTags } = g.Ver['Versions'];
  const startRow = rowTags.tablestart;
  const endRow = rowTags.tableend;

  if (startRow === undefined || endRow === undefined) {
    throw new Error("Could not find 'tablestart' or 'tableend' row tags in the master <Versions> sheet.");
  }

  for (let r = startRow; r <= endRow; r++) {
    const rowVersion = String(arr[r][colTags.version]);
    const rowAbbr = arr[r][colTags.ssabbr];

    if (rowVersion === version && rowAbbr === ssAbbr) {
      return arr[r][colTags.ssid]; // Return the ID as soon as we find the match
    }
  }

  return null; // Return null if no match is found after checking all rows
} // End function fGetMasterSheetId


/* function fGetSheetId
   Purpose: Gets a specific spreadsheet ID from the player's local collection, using a session cache-first approach.
   Assumptions: The Codex has a <MyVersions> sheet that has been populated.
   Notes: This is the primary function for retrieving any local versioned sheet ID.
   @param {string} version - The version number of the sheet ID to retrieve (e.g., '3').
   @param {string} ssAbbr - The abbreviated name of the sheet (e.g., 'CS').
   @returns {string} The spreadsheet ID.
*/
function fGetSheetId(version, ssAbbr) {
  // 1. Check if the in-memory session cache (g.sheetIDs) is empty.
  // If it is, load it from the spreadsheet. This happens once per session.
  if (Object.keys(g.sheetIDs).length === 0) {
    fLoadSheetIDsFromMyVersions();
  }

  // 2. Attempt to retrieve the ID from the now-populated session cache.
  if (g.sheetIDs[version] && g.sheetIDs[version][ssAbbr]) {
    return g.sheetIDs[version][ssAbbr].ssid;
  } else {
    // 3. If it's still not found, throw a clear error.
    throw new Error(`Could not find a local Sheet ID for version "${version}", abbreviation "${ssAbbr}". Check the <MyVersions> sheet.`);
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
  const codexSS = fGetCodexSpreadsheet(); // <-- THIS IS THE FIX

  // We now explicitly load from the Codex spreadsheet object.
  fLoadSheetToArray(ssKey, sheetName, codexSS);
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

