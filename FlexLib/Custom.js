/* global g, fGetSheetData, SpreadsheetApp, fPromptWithInput, fShowToast, fEndToast, fShowMessage, fGetCodexSpreadsheet, DriveApp, MailApp, Session */
/* exported fAddOwnCustomAbilitiesSource, fShareMyAbilities, fAddNewCustomSource */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Custom Abilities Management
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


/* function fShareMyAbilities
   Purpose: Orchestrates the workflow for a player to share their custom abilities sheet with another player.
   Assumptions: Run from the menu in a player's "Cust" sheet.
   Notes: Grants viewer permission and sends a notification email.
   @returns {void}
*/
function fShareMyAbilities() {
  fShowToast('⏳ Initializing share...', 'Share My Abilities');

  const ui = SpreadsheetApp.getUi();
  const activeSS = SpreadsheetApp.getActiveSpreadsheet();
  const sheetId = activeSS.getId();
  const sheetName = activeSS.getName();
  const currentUser = Session.getActiveUser().getEmail();

  // 1. Prompt for the recipient's email address
  const promptMessage = `Your Custom Abilities ID is:\n${sheetId}\n\nEnter the email address of the player you want to share this file with:`;
  const email = fPromptWithInput('Share My Abilities', promptMessage);

  if (!email) {
    fEndToast();
    fShowMessage('ℹ️ Canceled', 'Sharing was canceled.');
    return;
  }

  // 2. Basic email validation
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailRegex.test(email)) {
    fEndToast();
    fShowMessage('❌ Invalid Email', 'The email address you entered does not appear to be valid. Please try again.');
    return;
  }

  // 3. Grant permissions and send email
  try {
    fShowToast(`Sharing with ${email}...`, 'Share My Abilities');
    const file = DriveApp.getFileById(sheetId);
    file.addViewer(email); // <-- CHANGED to addViewer

    fShowToast('Sending notification email...', 'Share My Abilities');
    const subject = `Flex TTRPG: A custom abilities file has been shared with you!`;
    // --- NEW EMAIL BODY ---
    const body = `${currentUser} has shared a set of Flex Custom Abilities with you named "${sheetName}".\n\n` +
      `Please copy the ID string below EXACTLY as it appears and then on your Player's Codex, use the Flex menu's "Manage Custom Sources" and "Add New Source..." to paste this ID into.\n\n` +
      `ID String:\n${sheetId}`;
    MailApp.sendEmail(email, subject, body);

    fEndToast();
    fShowMessage('✅ Success!', `Your custom abilities file has been successfully shared with ${email}.`);
  } catch (e) {
    console.error(`Sharing failed. Error: ${e}`);
    fEndToast();
    fShowMessage('❌ Error', 'An error occurred while trying to share the file. Please ensure you are the owner of this file and try again.');
  }
} // End function fShareMyAbilities

/* function fAddNewCustomSource
   Purpose: The master workflow for adding a new, external custom ability source to the Codex.
   Assumptions: Run from the Codex menu.
   Notes: Includes validation, permission checks, and user prompts.
   @returns {void}
*/
function fAddNewCustomSource() {
  fShowToast('⏳ Initializing...', 'Add New Source');

  // 1. Prompt for the Sheet ID
  const sourceId = fPromptWithInput('Add Custom Source', 'Please enter the Google Sheet ID of the custom abilities file you want to add:');
  if (!sourceId) {
    fEndToast();
    fShowMessage('ℹ️ Canceled', 'Operation was canceled.');
    return;
  }

  // 2. Verify the ID and permissions
  let sourceSS;
  let ownerEmail;
  try {
    fShowToast('Verifying ID and permissions...', 'Add New Source');
    sourceSS = SpreadsheetApp.openById(sourceId);
    ownerEmail = sourceSS.getOwner() ? sourceSS.getOwner().getEmail() : 'Unknown';
  } catch (e) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not access the spreadsheet. Please check that the ID is correct and that the owner has shared the file with you.');
    return;
  }

  // 3. Check for duplicates
  const ssKey = 'Codex';
  const sheetName = 'CustomSources';
  const codexSS = fGetCodexSpreadsheet();
  const { arr, rowTags, colTags } = fGetSheetData(ssKey, sheetName, codexSS, true);
  const headerRow = rowTags.header;
  const sheetIdCol = colTags.sheetid;

  for (let r = headerRow + 1; r < arr.length; r++) {
    if (arr[r][sheetIdCol] === sourceId) {
      fEndToast();
      fShowMessage('⚠️ Duplicate', 'This custom source has already been added to your Codex.');
      return;
    }
  }

  // 4. Prompt for a friendly name with the updated example text
  const sourceName = fPromptWithInput('Name the Source', `✅ Success! File access verified.\n\nOwner: ${ownerEmail}\n\nPlease enter a friendly name for this source (e.g., "John's Custom List"):`);
  if (!sourceName) {
    fEndToast();
    fShowMessage('ℹ️ Canceled', 'Operation was canceled.');
    return;
  }

  // 5. Add the new source to the sheet
  fShowToast('Adding new source to your Codex...', 'Add New Source');
  const destSheet = codexSS.getSheetByName(sheetName);
  const lastRow = destSheet.getLastRow();
  let targetRow;

  const dataToWrite = [];
  dataToWrite[colTags.sheetid - 1] = sourceId;
  dataToWrite[colTags.sourcename - 1] = sourceName;
  dataToWrite[colTags.owner - 1] = ownerEmail;

  const firstDataRowIndex = headerRow + 1;
  const templateRow = firstDataRowIndex + 1; // 1-based template row
  const sourceNameCol = colTags.sourcename;

  if (arr.length <= firstDataRowIndex || !arr[firstDataRowIndex][sourceNameCol]) {
    targetRow = templateRow;
  } else {
    targetRow = lastRow + 1;
    destSheet.insertRowsAfter(lastRow, 1);
    const formatSourceRange = destSheet.getRange(templateRow, 1, 1, destSheet.getMaxColumns());
    const formatDestRange = destSheet.getRange(targetRow, 1, 1, destSheet.getMaxColumns());
    formatSourceRange.copyTo(formatDestRange, { formatOnly: true });
  }

  const targetRange = destSheet.getRange(targetRow, 2, 1, dataToWrite.length);
  targetRange.setValues([dataToWrite]);

  fEndToast();
  fShowMessage('✅ Success', `The custom source "${sourceName}" has been successfully added to your Codex.`);
} // End function fAddNewCustomSource

/* function fAddOwnCustomAbilitiesSource
   Purpose: Automatically finds the player's own 'Cust' file and logs it as the first entry in <CustomSources>.
   Assumptions: This is run at the end of the initial setup, so the <MyVersions> sheet is populated.
   Notes: Ensures the player always has access to their own custom content.
   @returns {void}
*/
function fAddOwnCustomAbilitiesSource() {
  const codexSS = fGetCodexSpreadsheet();

  // 1. Find the player's 'Cust' file ID from their local <MyVersions> sheet.
  const custId = fGetSheetId(g.CURRENT_VERSION, 'Cust');

  // 2. Get the destination sheet and its properties.
  const destSheetName = 'CustomSources';
  const destSS = SpreadsheetApp.getActiveSpreadsheet();
  const destSheet = destSS.getSheetByName(destSheetName);
  const { arr, rowTags, colTags } = fGetSheetData('Codex', destSheetName, codexSS, true);
  const headerRow = rowTags.header;
  const firstDataRow = headerRow + 2; // 1-based row number for the first data entry

  // 3. Prepare the data to be written.
  const dataToWrite = [];
  dataToWrite[colTags.sheetid - 1] = custId;
  dataToWrite[colTags.sourcename - 1] = 'My Custom Abilities';
  dataToWrite[colTags.owner - 1] = 'Me';

  // 4. Write the data to the first pre-formatted row.
  const targetRange = destSheet.getRange(firstDataRow, 2, 1, dataToWrite.length);
  targetRange.setValues([dataToWrite]);
} // End function fAddOwnCustomAbilitiesSource