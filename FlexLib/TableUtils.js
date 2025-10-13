/* global fBuildTagMaps, g */
/* exported fDeleteTableRow */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Table Management Utilities
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fDeleteTableRow
   Purpose: Deletes a row from a tagged table using Header-based logic.
   Assumptions: The sheet has a table with a 'Header' row tag. The rowNum is 1-based.
   Notes: This is the master helper for safe row deletion. It now uses fGetSheetData for performance.
   @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet object containing the table.
   @param {number} rowNum - The 1-based number of the row to delete.
   @returns {string} The action taken: 'deleted' or 'cleared'.
*/
function fDeleteTableRow(sheet, rowNum) {
  const sheetName = sheet.getName();
  // Use fGetSheetData to get the rowTags map, ensuring we read the latest state
  const { rowTags } = fGetSheetData('Temp', sheetName, sheet.getParent(), true);
  
  // Convert 0-based index from rowTags to a 1-based row number
  const headerRow = rowTags.header !== undefined ? rowTags.header + 1 : -1;

  if (headerRow === -1) {
    console.error(`fDeleteTableRow could not find a 'Header' tag in sheet: ${sheetName}`);
    sheet.deleteRow(rowNum); // Fallback to a simple delete
    return 'deleted';
  }

  // Case 1: The table has only one data row (or is empty).
  // Clear content instead of deleting the row to preserve formatting.
  if (sheet.getLastRow() <= headerRow + 1) {
    const rangeToClear = sheet.getRange(rowNum, 2, 1, sheet.getLastColumn() - 1);
    rangeToClear.clearContent();
    // Also uncheck any checkbox in the row
    sheet.getRange(rowNum, 1, 1, sheet.getMaxColumns()).uncheck();
    return 'cleared';
  }

  // Default Case: There are multiple data rows. Safely delete the entire row.
  sheet.deleteRow(rowNum);
  return 'deleted';
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