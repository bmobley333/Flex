/* global SpreadsheetApp, fParseA1Notation, fShowMessage */
/* exported fToggleDesignerVisibility */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Designer Visibility Toggles
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fToggleDesignerVisibility
   Purpose: Toggles the visibility of all designer sheets, rows, and columns.
   Assumptions: A sheet named "Hide>" exists to serve as the state marker.
   Notes: This is the main orchestrator for the Show/Hide All feature.
   @returns {void}
*/
function fToggleDesignerVisibility() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  const hideMarkerSheet = ss.getSheetByName('Hide>');

  if (!hideMarkerSheet) {
    fShowMessage('❌ Error', 'Could not find the "Hide>" sheet to determine state.');
    return;
  }

  const shouldHide = !hideMarkerSheet.isSheetHidden();

  SpreadsheetApp.flush(); // Apply pending changes before proceeding

  if (shouldHide) {
    fHideAllElements(allSheets, hideMarkerSheet.getIndex());
    fShowMessage('✅ Success', 'All designer elements have been hidden.');
  } else {
    fUnhideAllElements(allSheets, hideMarkerSheet.getIndex());
    fShowMessage('✅ Success', 'All designer elements have been shown.');
  }
} // End function fToggleDesignerVisibility

/* function fHideAllElements
   Purpose: Hides all designated designer elements.
   Assumptions: Called by function fToggleDesignerVisibility.
   Notes: Iterates through sheets to hide them and their specified rows/cols.
   @param {GoogleAppsScript.Spreadsheet.Sheet[]} allSheets - An array of all sheets in the spreadsheet.
   @param {number} hideMarkerIndex - The 1-based index of the "Hide>" sheet.
   @returns {void}
*/
function fHideAllElements(allSheets, hideMarkerIndex) {
  allSheets.forEach((sheet, index) => {
    // Hide all sheets from the marker onwards
    if (index + 1 >= hideMarkerIndex) {
      sheet.hideSheet();
    }

    const note = sheet.getRange('A1').getNote();
    if (note.includes('Hide: ')) {
      const hideString = note.split('Hide: ')[1].split('\n')[0];
      const ranges = fParseA1Notation(hideString);
      ranges.rows.forEach(row => sheet.hideRows(row));
      ranges.cols.forEach(col => sheet.hideColumns(col));
    }
  });
} // End function fHideAllElements

/* function fUnhideAllElements
   Purpose: Unhides all designated designer elements.
   Assumptions: Called by function fToggleDesignerVisibility.
   Notes: Iterates through sheets to unhide them and their specified rows/cols.
   @param {GoogleAppsScript.Spreadsheet.Sheet[]} allSheets - An array of all sheets in the spreadsheet.
   @param {number} hideMarkerIndex - The 1-based index of the "Hide>" sheet.
   @returns {void}
*/
function fUnhideAllElements(allSheets, hideMarkerIndex) {
  allSheets.forEach((sheet, index) => {
    // Unhide all sheets from the marker onwards
    if (index + 1 >= hideMarkerIndex) {
      sheet.showSheet();
    }

    const note = sheet.getRange('A1').getNote();
    if (note.includes('Hide: ')) {
      const hideString = note.split('Hide: ')[1].split('\n')[0];
      const ranges = fParseA1Notation(hideString);
      ranges.rows.forEach(row => sheet.unhideRow(sheet.getRange(row, 1)));
      ranges.cols.forEach(col => sheet.unhideColumn(sheet.getRange(1, col)));
    }
  });

  // Scroll to the top-left of the active sheet
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange('A1').activate();
} // End function fUnhideAllElements