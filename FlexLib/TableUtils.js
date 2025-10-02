/* global fBuildTagMaps, g */
/* exported fDeleteTableRow */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Table Management Utilities
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fDeleteTableRow
   Purpose: Deletes a row from a tagged table, correctly handling the transfer of TableStart/TableEnd tags.
   Assumptions: The sheet has a table with 'TableStart' and 'TableEnd' row tags. The rowNum is 1-based.
   Notes: This is the master helper for safe row deletion from tagged tables.
   @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet object containing the table.
   @param {number} rowNum - The 1-based number of the row to delete.
   @returns {void}
*/
function fDeleteTableRow(sheet, rowNum) {
  // Get the current state of the table before any action
  const sheetName = sheet.getName();
  fBuildTagMaps('Codex', sheetName);
  const { rowTags } = g.Codex[sheetName];
  const startRow = rowTags.tablestart + 1; // Convert to 1-based
  const endRow = rowTags.tableend + 1;   // Convert to 1-based

  // Case 1: Deleting the ONLY row in the table
  if (rowNum === startRow && rowNum === endRow) {
    // Do not delete the row, just clear its contents after the tag column
    const rangeToClear = sheet.getRange(rowNum, 2, 1, sheet.getLastColumn() - 1);
    rangeToClear.clearContent();
    return; // Stop further execution
  }

  // Case 2: Deleting the TableStart row
  if (rowNum === startRow) {
    const nextRowCell = sheet.getRange(rowNum + 1, 1);
    const currentRowCell = sheet.getRange(rowNum, 1);
    const combinedTags = fCleanTags(currentRowCell.getValue(), nextRowCell.getValue());
    nextRowCell.setValue(combinedTags);
    sheet.deleteRow(rowNum);
    return;
  }

  // Case 3: Deleting the TableEnd row
  if (rowNum === endRow) {
    const prevRowCell = sheet.getRange(rowNum - 1, 1);
    const currentRowCell = sheet.getRange(rowNum, 1);
    const combinedTags = fCleanTags(currentRowCell.getValue(), prevRowCell.getValue());
    prevRowCell.setValue(combinedTags);
    sheet.deleteRow(rowNum);
    return;
  }
  
  // Default Case: Deleting a middle row (no tag changes needed)
  sheet.deleteRow(rowNum);

} // End function fDeleteTableRow


/* function fCleanTags
   Purpose: Merges two tag strings into a single, clean, comma-separated string.
   Assumptions: None.
   Notes: Removes duplicates, extra spaces, and handles empty strings.
   @param {string} tagString1 - The first string of tags.
   @param {string} tagString2 - The second string of tags.
   @returns {string} The cleaned and merged tag string.
*/
function fCleanTags(tagString1, tagString2) {
  const combined = `${tagString1},${tagString2}`;
  const tags = combined.split(',')
    .map(tag => tag.trim()) // Remove leading/trailing spaces
    .filter(tag => tag);     // Remove any empty strings
  
  // Return a unique, sorted list of tags
  return [...new Set(tags)].sort().join(',');
} // End function fCleanTags