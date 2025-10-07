/* global SpreadsheetApp, fShowMessage, fParseA1Notation */
/* exported fPrepGameForPaper */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Paper & Print Utilities
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fPrepGameForPaper
   Purpose: Creates a clean, print-friendly copy of the <Game> sheet by duplicating it and removing hidden designer rows/columns.
   Assumptions: Run from a Character Sheet. A sheet named 'Game' and 'Paper' must exist. The 'Game' sheet must have a 'Hide:' note in cell A1.
   Notes: This is a designer-only function to prepare a sheet for printing or clean viewing.
   @returns {void}
*/
function fPrepGameForPaper() {
  fShowToast('⏳ Preparing <Paper> sheet...', 'Print Prep');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('Game');
  const oldPaperSheet = ss.getSheetByName('Paper');

  if (!sourceSheet) {
    fEndToast();
    fShowMessage('❌ Error', "A sheet named 'Game' must exist to run this function.");
    return;
  }

  // 1. Delete the old <Paper> sheet if it exists.
  if (oldPaperSheet) {
    ss.deleteSheet(oldPaperSheet);
  }

  // 2. Duplicate the <Game> sheet to create a perfect copy.
  const newPaperSheet = sourceSheet.copyTo(ss);
  newPaperSheet.setName('Paper');

  // 3. Move the new sheet to the very first position.
  ss.setActiveSheet(newPaperSheet);
  ss.moveActiveSheet(1);

  // 4. Read the 'Hide:' notation from the A1 cell.
  const note = newPaperSheet.getRange('A1').getNote();
  if (!note.includes('Hide: ')) {
    fEndToast();
    fShowMessage('✅ Success', 'The <Paper> sheet has been created, but no "Hide:" notation was found to remove designer elements.');
    return;
  }
  const hideString = note.split('Hide: ')[1].split('\n')[0];
  const rangesToHide = fParseA1Notation(hideString);

  // 5. Delete the specified columns and rows, starting from the end to avoid shifting indices.
  fShowToast('Removing designer elements...', 'Print Prep');
  rangesToHide.cols.sort((a, b) => b - a).forEach(col => newPaperSheet.deleteColumn(col));
  rangesToHide.rows.sort((a, b) => b - a).forEach(row => newPaperSheet.deleteRow(row));

  newPaperSheet.activate(); // Make it the active sheet.
  fEndToast();
  fShowMessage('✅ Success', 'The <Paper> sheet has been successfully created and cleaned for printing.');
} // End function fPrepGameForPaper