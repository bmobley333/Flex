/* global fShowMessage, fBuildTagMaps, g, fPromptWithInput, fGetSheetId, fGetOrCreateFolder, fSyncVersionFiles, DriveApp, SpreadsheetApp, fCreateNewCharacterSheet */
/* exported fCreateCharacter */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Character Management
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fCreateNewCharacterSheet
   Purpose: Creates and names a new character sheet from the local master and logs it in the Codex.
   Assumptions: The required master files for the selected version have already been synced and logged in <MyVersions>.
   Notes: This is the final step in the character creation workflow.
   @param {string} version - The game version for the new character (e.g., '3').
   @param {GoogleAppsScript.Drive.Folder} parentFolder - The "MetaScape Flex" folder object.
   @returns {void}
*/
function fCreateNewCharacterSheet(version, parentFolder) {
  // 1. Get the local CS template ID using the new ID Management system
  const localCsId = fGetSheetId(version, 'CS');

  if (!localCsId) {
    // This is a fallback error; fGetSheetId should throw its own specific error first.
    fShowMessage('❌ Error', `Could not find the local master Character Sheet for Version ${version}. Please try syncing versions again.`);
    return;
  }

  // 2. Copy the template and prompt for a name
  fShowMessage('New Character', '⏳ Creating a new character sheet...');
  const csTemplateFile = DriveApp.getFileById(localCsId);
  const newCharSheet = csTemplateFile.makeCopy(parentFolder);

  const characterName = fPromptWithInput('Name Your Character', 'Please enter a name for your new character:');

  if (!characterName) {
    newCharSheet.setTrashed(true);
    fShowMessage('ℹ️ Canceled', 'Character creation has been canceled.');
    return;
  }

  newCharSheet.setName(characterName);

  // 3. Log the new character in the Codex's <Characters> sheet
  const ssKey = 'Codex';
  const sheetName = 'Characters';
  fLoadSheetToArray(ssKey, sheetName);
  fBuildTagMaps(ssKey, sheetName);

  const destSS = SpreadsheetApp.getActiveSpreadsheet();
  const destSheet = destSS.getSheetByName(sheetName);
  const { arr, rowTags, colTags } = g[ssKey][sheetName];

  const startRow = rowTags.tablestart;
  const endRow = rowTags.tableend;
  const charNameCol = colTags.charname;

  let targetRow;

  // Prepare the character data first, as it's needed in both cases.
  const dataToWrite = [];
  dataToWrite[colTags.csid - 1] = newCharSheet.getId();
  dataToWrite[colTags.version - 1] = version;
  dataToWrite[colTags.checkbox - 1] = true;
  dataToWrite[colTags.charname - 1] = characterName; // Placeholder for rich text
  dataToWrite[colTags.rules - 1] = 'Rules'; // Placeholder for rich text


  // Case 1: First character, table is empty.
  if (startRow === endRow && (!arr[startRow] || arr[startRow][charNameCol] === '')) {
    targetRow = startRow + 1;
    // Data is written starting from the second column.
    const targetRange = destSheet.getRange(targetRow, 2, 1, dataToWrite.length);
    targetRange.setValues([dataToWrite]);
  } else {
    // Case 2 & 3: One or more characters exist.
    targetRow = endRow + 2;
    destSheet.insertRowsAfter(endRow + 1, 1);

    // Move the 'TableEnd' tag
    const oldTagCell = destSheet.getRange(endRow + 1, 1);
    const oldTags = oldTagCell.getValue().toString().split(',').map(t => t.trim());
    const newTags = oldTags.filter(t => t.toLowerCase() !== 'tableend');
    oldTagCell.setValue(newTags.join(', '));
    destSheet.getRange(targetRow, 1).setValue('TableEnd'); // Set the tag in the new row

    // Clear previous checkboxes
    if (colTags.checkbox !== undefined) {
      const checkboxCol = colTags.checkbox + 1;
      const numRows = endRow - startRow + 1;
      destSheet.getRange(startRow + 1, checkboxCol, numRows, 1).uncheck();
    }
    // Write the data starting from the second column, preserving the new tag
    const targetRange = destSheet.getRange(targetRow, 2, 1, dataToWrite.length);
    targetRange.setValues([dataToWrite]);
  }

  // 4. Format the new row appropriately
  // Set the checkbox data validation
  if (colTags.checkbox !== undefined) {
    destSheet.getRange(targetRow, colTags.checkbox + 1).insertCheckboxes();
  }
  // Set the rich text link for the character name
  const link = SpreadsheetApp.newRichTextValue().setText(characterName).setLinkUrl(newCharSheet.getUrl()).build();
  destSheet.getRange(targetRow, colTags.charname + 1).setRichTextValue(link);

  // Set the rich text link for the Rules document
  const rulesId = fGetSheetId(version, 'Rules');
  const rulesUrl = `https://docs.google.com/document/d/${rulesId}/`;
  const rulesLink = SpreadsheetApp.newRichTextValue().setText('Rules').setLinkUrl(rulesUrl).build();
  destSheet.getRange(targetRow, colTags.rules + 1).setRichTextValue(rulesLink);


  // 5. Final, corrected success message
  const successMessage = `✅ Success! Your new character, "${characterName}," has been created.\n\nA link has been added to your <Characters> sheet.`;
  fShowMessage('Character Created!', successMessage);
} // End function fCreateNewCharacterSheet


/* function fCreateCharacter
   Purpose: The master orchestrator for the entire character creation workflow.
   Assumptions: The Codex has a <MyVersions> sheet.
   Notes: Intelligently triggers a one-time setup if needed, otherwise proceeds to character creation.
   @returns {void}
*/
function fCreateCharacter() {
  const ssKey = 'Codex';
  const sheetName = 'MyVersions';

  // 1. Check if the one-time setup is needed.
  fLoadSheetToArray(ssKey, sheetName);
  fBuildTagMaps(ssKey, sheetName);

  let { arr, rowTags, colTags } = g[ssKey][sheetName];
  const startRow = rowTags.tablestart;
  const endRow = rowTags.tableend;
  const ssAbbrCol = colTags.ssabbr;

  // Condition for an empty table (first-time use)
  if (startRow === endRow && (!arr[startRow] || arr[startRow][ssAbbrCol] === '')) {
    fInitialSetup();
    // After setup, we MUST reload the sheet data to get the new information
    fLoadSheetToArray(ssKey, sheetName);
    fBuildTagMaps(ssKey, sheetName);
    // Re-assign our local variables with the new data
    let reloadedData = g[ssKey][sheetName];
    arr = reloadedData.arr;
    rowTags = reloadedData.rowTags;
    colTags = reloadedData.colTags;
  }

  // 2. Load available versions from the now-populated <MyVersions> sheet.
  const versionCol = colTags.version;
  const versionsStartRow = rowTags.tablestart;
  const versionsEndRow = rowTags.tableend;

  // Create a unique list of versions that have a 'CS' file available.
  const availableVersions = [...new Set(
    arr.slice(versionsStartRow, versionsEndRow + 1)
       .filter(row => row[colTags.ssabbr] === 'CS')
       .map(row => String(row[versionCol]))
  )];


  if (availableVersions.length === 0) {
    fShowMessage('❌ Error', 'No game versions with a Character Sheet (CS) were found in your <MyVersions> sheet.');
    return;
  }

  // 3. Prompt user for selection
  const promptMessage = `Please enter the game version you would like to use for your new character.\n\nAvailable versions:\n${availableVersions.join(', ')}`;
  const selectedVersion = fPromptWithInput('Select Game Version', promptMessage);

  // 4. Handle response
  if (selectedVersion === null) {
    fShowMessage('ℹ️ Canceled', 'Character creation has been canceled.');
    return;
  }

  if (!availableVersions.includes(selectedVersion)) {
    fShowMessage('❌ Error', `Invalid version selected. Please enter one of the available versions: ${availableVersions.join(', ')}`);
    return;
  }

  // 5. Create the new character sheet and log it.
  const parentFolder = fGetOrCreateFolder('MetaScape Flex'); // We still need the folder reference
  fCreateNewCharacterSheet(selectedVersion, parentFolder);

} // End function fCreateCharacter