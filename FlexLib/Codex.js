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
    console.error(`Could not embed Codex ID because the template is missing a <Data> sheet.`);
    return;
  }

  const codexId = fGetCodexSpreadsheet().getId();

  // --- THIS IS THE FIX ---
  // Use the architecturally correct gatekeeper function to get sheet data.
  const { rowTags, colTags } = fGetSheetData('Temp', 'Data', ss, true); // Force a refresh

  const rowIndex = rowTags.codexid;
  const colIndex = colTags.data;

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

  // 3. Try to find the Codex ID in the <Data> sheet using the correct gatekeeper.
  try {
    // --- THIS IS THE FIX ---
    // We pass activeSS directly to fGetSheetData to prevent circular lookups.
    const { arr, rowTags, colTags } = fGetSheetData('Temp', 'Data', activeSS, true);

    if (rowTags.codexid !== undefined && colTags.data !== undefined) {
      const codexId = arr[rowTags.codexid][colTags.data];
      if (codexId) {
        g.codexSS = SpreadsheetApp.openById(codexId);
        return g.codexSS;
      }
    }
  } catch (e) {
    console.error(`Could not read Codex ID from <Data> sheet. Assuming active sheet is the Codex. Error: ${e}`);
  }

  // 4. If all else fails, assume we ARE the Codex.
  g.codexSS = activeSS;
  return g.codexSS;

} // End function fGetCodexSpreadsheet