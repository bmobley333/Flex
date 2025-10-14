/* global SpreadsheetApp, fNormalizeTags, fShowMessage */
/* exported fVerifyActiveSheetTags */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Sheet Verification Utilities
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fVerifyActiveSheetTags
   Purpose: Verifies unique column and row tags on the currently active sheet by leveraging fGetSheetData.
   Assumptions: The function is triggered by a user on an active sheet.
   Notes: This version is faster and tests the same cached tag maps the rest of the system uses.
   @returns {void}
*/
function fVerifyActiveSheetTags() {
  fShowToast('⏳ Verifying all tags...', 'Tag Verification');
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();

  try {
    // Use fGetSheetData to get the canonical tag maps for this sheet, forcing a refresh
    const { arr } = fGetSheetData('Temp', sheetName, SpreadsheetApp.getActiveSpreadsheet(), true);

    // 1. Verify Column Tags
    const seenColTags = {};
    const colTagRow = arr[0] || [];
    for (let c = 0; c < colTagRow.length; c++) {
      const normalizedTags = fNormalizeTags(colTagRow[c]);
      for (const tag of normalizedTags) {
        if (seenColTags[tag]) {
          const message = `Duplicate column tag found: "${tag}"\n\nOriginal in cell: ${seenColTags[tag]}\nDuplicate in cell: ${sheet.getRange(1, c + 1).getA1Notation()}`;
          fEndToast();
          fShowMessage('⚠️ Tag Verification Failed', message);
          return;
        }
        seenColTags[tag] = sheet.getRange(1, c + 1).getA1Notation();
      }
    }

    // 2. Verify Row Tags
    const seenRowTags = {};
    for (let r = 0; r < arr.length; r++) {
      // Ensure the row and cell exist before trying to read from it
      if (arr[r] && arr[r][0]) {
        const normalizedTags = fNormalizeTags(arr[r][0]);
        for (const tag of normalizedTags) {
          if (seenRowTags[tag]) {
            const message = `Duplicate row tag found: "${tag}"\n\nOriginal in cell: ${seenRowTags[tag]}\nDuplicate in cell: ${sheet.getRange(r + 1, 1).getA1Notation()}`;
            fEndToast();
            fShowMessage('⚠️ Tag Verification Failed', message);
            return;
          }
          seenRowTags[tag] = sheet.getRange(r + 1, 1).getA1Notation();
        }
      }
    }

    fEndToast();
    fShowMessage('✅ Tag Verification', '✅ Success! All column and row tags are unique.');

  } catch (e) {
    fEndToast();
    fShowMessage('❌ Error', `An error occurred during verification: ${e.message}`);
  }
} // End function fVerifyActiveSheetTags