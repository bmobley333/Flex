/* global SpreadsheetApp, fNormalizeTags, fShowMessage */
/* exported fVerifyActiveSheetTags */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Sheet Verification Utilities
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fVerifyActiveSheetTags
   Purpose: Verifies unique column and row tags on the currently active sheet.
   Assumptions: The function is triggered by a user on an active sheet.
   Notes: Stops and reports the first duplicate found.
   @returns {void}
*/
function fVerifyActiveSheetTags() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();

  // 1. Verify Column Tags (Row 0)
  const seenColTags = {};
  const colTags = data[0] || []; // Handle empty sheets
  for (let c = 0; c < colTags.length; c++) {
    const normalizedTags = fNormalizeTags(colTags[c]);
    for (const tag of normalizedTags) {
      const locationA1 = sheet.getRange(1, c + 1).getA1Notation();
      if (seenColTags[tag]) {
        const message = `Duplicate column tag found: "${tag}"\n\nOriginal: ${seenColTags[tag]}\nDuplicate: ${locationA1}`;
        fShowMessage('⚠️ Tag Verification Failed', message);
        return;
      }
      seenColTags[tag] = locationA1;
    }
  }

  // 2. Verify Row Tags (Column 0)
  const seenRowTags = {};
  for (let r = 0; r < data.length; r++) {
    const normalizedTags = fNormalizeTags(data[r][0]);
    for (const tag of normalizedTags) {
      const locationA1 = sheet.getRange(r + 1, 1).getA1Notation();
      if (seenRowTags[tag]) {
        const message = `Duplicate row tag found: "${tag}"\n\nOriginal: ${seenRowTags[tag]}\nDuplicate: ${locationA1}`;
        fShowMessage('⚠️ Tag Verification Failed', message);
        return;
      }
      seenRowTags[tag] = locationA1;
    }
  }

  fShowMessage('Tag Verification', '✅ Success! All column and row tags are unique.');
} // End function fVerifyActiveSheetTags