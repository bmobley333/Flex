/* global SpreadsheetApp, fParseA1Notation, fShowMessage */
/* exported fToggleDesignerVisibility */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Designer Visibility Toggles
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fToggleDesignerVisibility
   Purpose: Toggles the visibility of all designer sheets, rows, and columns.
   Assumptions: A sheet named "Hide>" may exist to serve as a state marker.
   Notes: This is the main orchestrator for the Show/Hide All feature. It now handles cases where "Hide>" is missing.
   @returns {void}
*/
function fToggleDesignerVisibility() {
  fShowToast('⏳ Toggling visibility...', 'Show/Hide All');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  const hideMarkerSheet = ss.getSheetByName('Hide>');
  let shouldHide;

  if (hideMarkerSheet) {
    shouldHide = !hideMarkerSheet.isSheetHidden();
  } else {
    // Fallback: Determine state by checking the first available element
    for (const sheet of allSheets) {
      const note = sheet.getRange('A1').getNote();
      if (note.includes('Hide: ')) {
        const hideString = note.split('Hide: ')[1].split('\n')[0];
        const ranges = fParseA1Notation(hideString);
        if (ranges.rows.length > 0) {
          shouldHide = !sheet.isRowHiddenByUser(ranges.rows[0]);
        } else if (ranges.cols.length > 0) {
          shouldHide = !sheet.isColumnHiddenByUser(ranges.cols[0]);
        }
        break; // State determined, no need to check other sheets
      }
    }
  }

  if (shouldHide === undefined) {
    fEndToast();
    fShowMessage('ℹ️ No Action', 'No designer elements with "Hide:" notes were found to toggle.');
    return;
  }

  SpreadsheetApp.flush(); // Apply pending changes before proceeding

  const hideMarkerIndex = hideMarkerSheet ? hideMarkerSheet.getIndex() : -1;

  if (shouldHide) {
    fHideAllElements(allSheets, hideMarkerIndex);
    fEndToast();
    fShowMessage('✅ Success', 'All designer elements have been hidden.');
  } else {
    fUnhideAllElements(allSheets, hideMarkerIndex);
    fEndToast();
    fShowMessage('✅ Success', 'All designer elements have been shown.');
  }
} // End function fToggleDesignerVisibility

/* function fHideAllElements
   Purpose: Hides all designated designer elements.
   Assumptions: Called by  fToggleDesignerVisibility.
   Notes: Iterates through sheets to hide them and their specified rows/cols.
   @param {GoogleAppsScript.Spreadsheet.Sheet[]} allSheets - An array of all sheets in the spreadsheet.
   @param {number} hideMarkerIndex - The 1-based index of the "Hide>" sheet, or -1 if not found.
   @returns {void}
*/
function fHideAllElements(allSheets, hideMarkerIndex) {
  allSheets.forEach((sheet, index) => {
    const currentIndex = index + 1;
    // Only hide the sheet if a marker exists and we are at or after it
    if (hideMarkerIndex !== -1 && currentIndex >= hideMarkerIndex) {
      sheet.hideSheet();
    }

    // Only process rows/cols for sheets BEFORE the marker (or all sheets if no marker)
    if (hideMarkerIndex === -1 || currentIndex < hideMarkerIndex) {
      const note = sheet.getRange('A1').getNote();
      if (note.includes('Hide: ')) {
        const hideString = note.split('Hide: ')[1].split('\n')[0];
        const ranges = fParseA1Notation(hideString);
        ranges.rows.forEach(row => sheet.hideRows(row));
        ranges.cols.forEach(col => sheet.hideColumns(col));
      }
    }
  });
} // End function fHideAllElements

/* function fUnhideAllElements
   Purpose: Unhides all designated designer elements.
   Assumptions: Called by function fToggleDesignerVisibility.
   Notes: Iterates through sheets to unhide them and their specified rows/cols.
   @param {GoogleAppsScript.Spreadsheet.Sheet[]} allSheets - An array of all sheets in the spreadsheet.
   @param {number} hideMarkerIndex - The 1-based index of the "Hide>" sheet, or -1 if not found.
   @returns {void}
*/
function fUnhideAllElements(allSheets, hideMarkerIndex) {
  allSheets.forEach((sheet, index) => {
    const currentIndex = index + 1;
    // Only unhide the sheet if a marker exists and we are at or after it
    if (hideMarkerIndex !== -1 && currentIndex >= hideMarkerIndex) {
      sheet.showSheet();
    }

    // Only process rows/cols for sheets BEFORE the marker (or all sheets if no marker)
    if (hideMarkerIndex === -1 || currentIndex < hideMarkerIndex) {
      const note = sheet.getRange('A1').getNote();
      if (note.includes('Hide: ')) {
        const hideString = note.split('Hide: ')[1].split('\n')[0];
        const ranges = fParseA1Notation(hideString);
        ranges.rows.forEach(row => sheet.unhideRow(sheet.getRange(row, 1)));
        ranges.cols.forEach(col => sheet.unhideColumn(sheet.getRange(1, col)));
      }
    }
  });

  // Scroll to the top-left of the active sheet
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange('A1').activate();
} // End function fUnhideAllElements