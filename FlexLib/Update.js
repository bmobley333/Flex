/* global g, fBuildTagMaps, SpreadsheetApp, fShowMessage, PropertiesService */
/* exported fGetLatestVersions */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Update & Sync Logic
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fGetLatestVersions
   Purpose: Fetches the latest version data from the master Ver sheet and updates the local Codex.
   Assumptions: The master Ver ID is correctly set in g.MASTER_VER_ID.
   Notes: This performs a full overwrite of the local version data and rebuilds the cache.
   @returns {void}
*/
function fGetLatestVersions() {
  // 1. Get Source Data from Master Ver Sheet
  const sourceSS = SpreadsheetApp.openById(g.MASTER_VER_ID);
  const sourceSheet = sourceSS.getSheetByName('Versions');
  const sourceData = sourceSheet.getDataRange().getValues();
  const sourceRowTags = {};
  const sourceColTags = {};

  // Build source tag maps manually
  sourceData[0].forEach((tag, c) => fNormalizeTags(tag).forEach(t => (sourceColTags[t] = c)));
  sourceData.forEach((row, r) => fNormalizeTags(row[0]).forEach(t => (sourceRowTags[t] = r)));

  const sourceStart = sourceRowTags.tablestart;
  const sourceEnd = sourceRowTags.tableend;
  const sourceColsToGet = ['version', 'releasedate', 'ismaster', 'ssfullname', 'ssabbr', 'ssid'].map(t => sourceColTags[t]);
  const dataToPaste = sourceData.slice(sourceStart, sourceEnd + 1).map(row => sourceColsToGet.map(c => row[c]));

  // 2. Prepare Destination Sheet (Codex)
  const destSS = SpreadsheetApp.getActiveSpreadsheet();
  const destSheet = destSS.getSheetByName('Versions');
  fBuildTagMaps('Codex', 'Versions');
  const { rowTags, colTags } = g.Codex.Versions;
  const destStart = rowTags.tablestart;
  const destEnd = rowTags.tableend;
  const destNumRows = destEnd - destStart + 1;
  const destColsToSet = ['version', 'releasedate', 'ismaster', 'ssfullname', 'ssabbr', 'ssid'].map(t => colTags[t]);

  // 3. Clear old data and adjust rows
  if (destNumRows > 0) {
    destSheet.getRange(destStart + 1, 2, destNumRows, destSheet.getLastColumn() - 1).clearContent();
  }

  const rowDiff = dataToPaste.length - destNumRows;
  if (rowDiff > 0) {
    destSheet.insertRowsAfter(destEnd, rowDiff);
  } else if (rowDiff < 0) {
    destSheet.deleteRows(destEnd + rowDiff + 1, -rowDiff);
  }

  // 4. Write new data
  if (dataToPaste.length > 0) {
    const pasteRange = destSheet.getRange(destStart + 1, destColsToSet[0] + 1, dataToPaste.length, dataToPaste[0].length);
    pasteRange.setValues(dataToPaste);
  }

  // 5. Finalize by clearing the old cache and immediately rebuilding it.
  PropertiesService.getScriptProperties().deleteProperty('sheetIDs');
  g.sheetIDs = {}; // Clear in-memory cache
  fLoadSheetIDsFromCodex(); // Re-load from the sheet we just updated
  fCacheSheetIDsToStorage(); // Save the fresh data to persistent storage

  fShowMessage('✅ Success', 'The latest version data has been successfully loaded and cached.');
} // End function fGetLatestVersions

/* function fNormalizeTags
   Purpose: Converts a raw tag string into an array of standardized tags.
   Assumptions: None.
   Notes: This is a temporary copy until we resolve file dependencies.
   @param {string} tagString - The raw string from a tag cell.
   @returns {string[]} An array of normalized tags.
*/
function fNormalizeTags(tagString) {
  if (!tagString || typeof tagString !== 'string') return [];
  return tagString
    .toLowerCase()
    .replace(/\s+/g, '')
    .split(',')
    .filter(tag => tag);
} // End function fNormalizeTags