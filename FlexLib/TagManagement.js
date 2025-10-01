/* global g, fLoadSheetToArray, fNormalizeTags */
/* exported fBuildTagMaps */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Tag Management
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fBuildTagMaps
   Purpose: Builds row and column tag maps for a specified sheet for easy data access.
   Assumptions: The specified sheet exists in the active spreadsheet.
   Notes: The maps are stored in the g object alongside the sheet's array data.
   @param {string} ssKey - The key for the spreadsheet in the g object (e.g., 'Codex').
   @param {string} sheetName - The exact, case-sensitive name of the sheet.
   @returns {void}
*/
function fBuildTagMaps(ssKey, sheetName) {
  // Ensure the sheet array exists, loading it if necessary.
  if (!g[ssKey] || !g[ssKey][sheetName] || !g[ssKey][sheetName].arr) {
    fLoadSheetToArray(ssKey, sheetName);
  }

  const data = g[ssKey][sheetName].arr;
  const rowTagMap = {};
  const colTagMap = {};

  // 1. Build Column Tag Map (from row 0)
  const colTags = data[0] || [];
  for (let c = 0; c < colTags.length; c++) {
    const normalizedTags = fNormalizeTags(colTags[c]);
    normalizedTags.forEach(tag => {
      colTagMap[tag] = c;
    });
  }

  // 2. Build Row Tag Map (from column 0)
  for (let r = 0; r < data.length; r++) {
    // Ensure the row and cell exist to prevent errors on sparse data
    if (data[r] && typeof data[r][0] !== 'undefined') {
      const normalizedTags = fNormalizeTags(data[r][0]);
      normalizedTags.forEach(tag => {
        rowTagMap[tag] = r;
      });
    }
  }

  // 3. Store the maps in the global object
  g[ssKey][sheetName].rowTags = rowTagMap;
  g[ssKey][sheetName].colTags = colTagMap;
} // End function fBuildTagMaps