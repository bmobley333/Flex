/* global fBuildTagMaps, g */
/* exported fDeleteTableRow */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Table Management Utilities
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fDeleteTableRow
   Purpose: Deletes a row from a tagged table using Header-based logic.
   Assumptions: The sheet has a table with a 'Header' row tag. The rowNum is 1-based.
   Notes: This is the master helper for safe row deletion. It now returns the action taken.
   @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet object containing the table.
   @param {number} rowNum - The 1-based number of the row to delete.
   @returns {string} The action taken: 'deleted' or 'cleared'.
*/
function fDeleteTableRow(sheet, rowNum) {
  const lastRow = sheet.getLastRow();
  const dataRange = sheet.getRange('A1:A').getValues();
  let headerRow = -1;
  for (let i = 0; i < dataRange.length; i++) {
    if (fNormalizeTags(dataRange[i][0]).includes('header')) {
      headerRow = i + 1; // 1-based
      break;
    }
  }

  if (headerRow === -1) {
    console.error(`fDeleteTableRow could not find a 'Header' tag in sheet: ${sheet.getName()}`);
    sheet.deleteRow(rowNum); // Fallback to a simple delete
    return 'deleted';
  }

  // Case 1: The table has only one data row (or is empty).
  // Instead of deleting the row (which removes formatting), we clear its content.
  if (lastRow <= headerRow + 1) {
    const rangeToClear = sheet.getRange(rowNum, 2, 1, sheet.getLastColumn() - 1);
    rangeToClear.clearContent();
    sheet.getRange(rowNum, 1, 1, sheet.getMaxColumns()).uncheck();
    return 'cleared'; // Return the 'cleared' status
  }

  // Default Case: There are multiple data rows. Safely delete the entire row.
  sheet.deleteRow(rowNum);
  return 'deleted'; // Return the 'deleted' status
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