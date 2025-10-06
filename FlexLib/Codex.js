/* global g, SpreadsheetApp, fBuildTagMaps, fLoadSheetToArray */
/* exported fGetCodexSpreadsheet */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Codex Management
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fEmbedCodexId
   Purpose: Ensures a <Data> sheet exists and writes the Codex ID into the correctly tagged cell.
   Assumptions: The spreadsheet file has a pre-configured <Data> sheet with a 'CodexID' row tag and a 'Data' column tag.
   Notes: This is the definitive helper for embedding the Codex ID into a file.
   @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - The spreadsheet object to embed the ID into.
   @returns {void}
*/
function fEmbedCodexId(ss) {
  const dataSheet = ss.getSheetByName('Data');
  if (!dataSheet) {
    // If the template is missing the sheet, we can't proceed.
    console.error(`Could not embed Codex ID because the template is missing a <Data> sheet.`);
    return;
  }

  const codexId = fGetCodexSpreadsheet().getId();

  // This is a rare case where we do a direct write without a full fGetSheetData
  // call, as we need to be very targeted to avoid cache-loop issues.
  const dataArr = dataSheet.getDataRange().getValues();
  const colTags = {};
  dataArr[0].forEach((tag, c) => fNormalizeTags(tag).forEach(t => (colTags[t] = c)));
  const rowTags = {};
  dataArr.forEach((row, r) => fNormalizeTags(row[0]).forEach(t => (rowTags[t] = r)));

  const rowIndex = rowTags.codexid;
  const colIndex = colTags.data; // <-- Updated from colTags.codexid

  if (rowIndex !== undefined && colIndex !== undefined) {
    dataSheet.getRange(rowIndex + 1, colIndex + 1).setValue(codexId);
  } else {
    console.error(`Could not embed Codex ID because the <Data> sheet is missing the 'CodexID' (row) or 'Data' (column) tags.`);
  }
} // End function fEmbedCodexId

/* function fGetCodexSpreadsheet
   Purpose: Gets the Spreadsheet object for the Player's Codex, creating a session-based cache for it.
   Assumptions: If not run from the Codex, the active sheet has a <Data> sheet with a cell tagged 'CodexID' (row) and 'Data' (column).
   Notes: This is the definitive helper for finding the Codex from any context.
   @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} The Spreadsheet object for the Player's Codex.
*/
function fGetCodexSpreadsheet() {
  // 1. Return the cached object if it already exists.
  if (g.codexSS) {
    return g.codexSS;
  }

  const activeSS = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = activeSS.getSheetByName('Data');

  // 2. If no <Data> sheet, assume we ARE the Codex.
  if (!dataSheet) {
    g.codexSS = activeSS;
    return g.codexSS;
  }

  // 3. Try to find the Codex ID in the <Data> sheet.
  try {
    // We use fLoadSheetToArray directly here to avoid circular dependencies
    const dataArr = dataSheet.getDataRange().getValues();
    const colTags = {};
    dataArr[0].forEach((tag, c) => fNormalizeTags(tag).forEach(t => (colTags[t] = c)));
    const rowTags = {};
    dataArr.forEach((row, r) => fNormalizeTags(row[0]).forEach(t => (rowTags[t] = r)));

    if (rowTags.codexid !== undefined && colTags.data !== undefined) {
      const codexId = dataArr[rowTags.codexid][colTags.data];
      if (codexId) {
        g.codexSS = SpreadsheetApp.openById(codexId);
        return g.codexSS;
      }
    }
  } catch (e) {
    // If any error occurs trying to read the tag, fall back to assuming we are the Codex.
    console.error(`Could not read Codex ID from <Data> sheet. Assuming active sheet is the Codex. Error: ${e}`);
  }

  // 4. If all else fails, assume we ARE the Codex.
  g.codexSS = activeSS;
  return g.codexSS;

} // End function fGetCodexSpreadsheet