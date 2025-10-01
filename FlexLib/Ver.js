/* global g, fLoadSheetToArray, fNormalizeTags, fShowMessage */
/* exported run */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Dispatcher & Ver Logic
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function run
   Purpose: Acts as the central dispatcher for all commands initiated from a local sheet script.
   Assumptions: The command string passed matches a key in the commandMap.
   Notes: This provides a single entry point and a master try/catch for robust error handling.
   @param {string} command - The unique identifier for the command to execute.
   @returns {void}
*/
function run(command) {
  try {
    const commandMap = {
      Ver_TagVerification: fVerVerifyTags,
      // Future commands will be added here
    };

    if (commandMap[command]) {
      commandMap[command]();
    } else {
      throw new Error(`Unknown command received: ${command}`);
    }
  } catch (e) {
    console.error(e);
    fShowMessage('❌ Error', e.message);
  }
} // End function run

/* function fVerVerifyTags
   Purpose: Verifies that all column tags (row 0) and row tags (col 0) are unique within their respective domains.
   Assumptions: The 'Versions' sheet exists in the active spreadsheet.
   Notes: Stops and reports the first duplicate found.
   @returns {void}
*/
function fVerVerifyTags() {
  const SPREADSHEET_KEY = 'Ver';
  const SHEET_NAME = 'Versions';

  fLoadSheetToArray(SPREADSHEET_KEY, SHEET_NAME);
  const data = g[SPREADSHEET_KEY][SHEET_NAME].arr;

  // 1. Verify Column Tags (Row 0)
  const seenColTags = {};
  const colTags = data[0];
  for (let c = 0; c < colTags.length; c++) {
    const normalizedTags = fNormalizeTags(colTags[c]);
    for (const tag of normalizedTags) {
      const locationA1 = String.fromCharCode(65 + c) + '1';
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
      const locationA1 = 'A' + (r + 1);
      if (seenRowTags[tag]) {
        const message = `Duplicate row tag found: "${tag}"\n\nOriginal: ${seenRowTags[tag]}\nDuplicate: ${locationA1}`;
        fShowMessage('⚠️ Tag Verification Failed', message);
        return;
      }
      seenRowTags[tag] = locationA1;
    }
  }

  fShowMessage('Tag Verification', '✅ Success! All column and row tags are unique.');
} // End function fVerVerifyTags