/* global g, fBuildTagMaps, fShowMessage */
/* exported fTestTagMaps */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Testing Utilities
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fTestTagMaps
   Purpose: A test function to verify that fBuildTagMaps is working correctly.
   Assumptions: The 'Codex' spreadsheet has a sheet named 'Versions' with tags.
   Notes: Displays the first and last found tags for rows and columns.
   @returns {void}
*/
function fTestTagMaps() {
  fBuildTagMaps('Codex', 'Versions');

  const { rowTags, colTags } = g.Codex.Versions;

  const rowKeys = Object.keys(rowTags);
  const colKeys = Object.keys(colTags);

  let message = 'ℹ️ Tag mapping successful!\n\n';
  message += `Row Tags Found: ${rowKeys.length}\n`;
  if (rowKeys.length > 0) {
    message += `First: "${rowKeys[0]}" -> index ${rowTags[rowKeys[0]]}\n`;
    message += `Last: "${rowKeys[rowKeys.length - 1]}" -> index ${rowTags[rowKeys[rowKeys.length - 1]]}\n`;
  }

  message += `\nCol Tags Found: ${colKeys.length}\n`;
  if (colKeys.length > 0) {
    message += `First: "${colKeys[0]}" -> index ${colTags[colKeys[0]]}\n`;
    message += `Last: "${colKeys[colKeys.length - 1]}" -> index ${colTags[colKeys[colKeys.length - 1]]}\n`;
  }

  fShowMessage('Test Results', message);
} // End function fTestTagMaps