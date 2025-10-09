/* global g, PropertiesService, SpreadsheetApp, fBuildTagMaps, fLoadSheetToArray */
/* exported fGetSheetId, fLoadSheetIDsFromMyVersions, fGetVerifiedLocalFile */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - ID Management & Caching
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


/* function fGetVerifiedLocalFile
   Purpose: A robust gatekeeper to get a local file object, self-healing if the file is missing.
   Assumptions: The file should exist in the player's <MyVersions> sheet.
   Notes: This is the primary function for safely accessing any local master file (CS, DB, etc.).
   @param {string} version - The version number of the file to retrieve.
   @param {string} ssAbbr - The abbreviated name of the sheet (e.g., 'CS', 'DB').
   @returns {GoogleAppsScript.Drive.File|null} The file object, or null if it cannot be found or restored.
*/
function fGetVerifiedLocalFile(version, ssAbbr) {
  const localId = fGetSheetId(version, ssAbbr);
  try {
    // --- Happy Path ---
    const file = DriveApp.getFileById(localId);
    return file;
  } catch (e) {
    // --- Self-Heal Path ---
    fShowToast(`⚠️ A core file (${ssAbbr}) is missing. Restoring...`, 'File Health Check');

    const masterCopiesFolder = fGetSubFolder('Master Copies - DO NOT DELETE');
    if (!masterCopiesFolder) {
      fEndToast();
      return null;
    }

    const masterId = fGetMasterSheetId(version, ssAbbr);
    if (!masterId) {
      fShowMessage('❌ Error', `Could not find the original master record for "${ssAbbr}" v${version} to restore it.`);
      return null;
    }

    const fileName = `v${version} MASTER_${ssAbbr} - DO NOT DELETE`;
    const newFile = DriveApp.getFileById(masterId).makeCopy(fileName, masterCopiesFolder);
    const newId = newFile.getId();

    const codexSS = fGetCodexSpreadsheet();
    const myVersionsSheet = codexSS.getSheetByName('MyVersions');
    const { arr, rowTags, colTags } = fGetSheetData('Codex', 'MyVersions', codexSS, true);
    const headerRow = rowTags.header;

    for (let r = headerRow + 1; r < arr.length; r++) {
      if (String(arr[r][colTags.version]) === version && arr[r][colTags.ssabbr] === ssAbbr) {
        myVersionsSheet.getRange(r + 1, colTags.ssid + 1).setValue(newId);
        break;
      }
    }

    fLoadSheetIDsFromMyVersions();
    fShowToast(`✅ Successfully restored ${ssAbbr}!`, 'File Health Check', 5);

    // --- NEW: Post-Heal Link Repair ---
    if (ssAbbr === 'Rules') {
      fUpdateCharacterRulesLinks(version, newId);
    }
    // --- END NEW ---

    return newFile;
  }
} // End function fGetVerifiedLocalFile

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
  const headerRow = rowTags.header;

  if (headerRow === undefined) {
    throw new Error('Could not find "Header" row tag in the master <Versions> sheet.');
  }
  const startRow = headerRow + 1;

  for (let r = startRow; r < arr.length; r++) {
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
   Assumptions: The <MyVersions> sheet is tagged with 'Header', 'version', 'ssabbr', and 'ssid'.
   Notes: This powers the cache with the player's own file data.
   @returns {void}
*/
function fLoadSheetIDsFromMyVersions() {
  const ssKey = 'Codex';
  const sheetName = 'MyVersions';
  const codexSS = fGetCodexSpreadsheet();

  fLoadSheetToArray(ssKey, sheetName, codexSS);
  fBuildTagMaps(ssKey, sheetName);

  const { arr, rowTags, colTags } = g[ssKey][sheetName];
  const headerRow = rowTags.header;

  if (headerRow === undefined) {
    throw new Error(`Could not find a "Header" tag in the <${sheetName}> sheet.`);
  }
  const startRow = headerRow + 1;

  // Clear the in-memory cache before reloading
  g.sheetIDs = {};

  for (let r = startRow; r < arr.length; r++) {
    // Ensure the row has data before trying to process it
    if (arr[r] && arr[r][colTags.ssabbr]) {
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
  }
} // End function fLoadSheetIDsFromMyVersions

