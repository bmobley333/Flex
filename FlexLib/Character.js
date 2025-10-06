/* global fShowMessage, fBuildTagMaps, g, fPromptWithInput, fGetSheetId, fGetOrCreateFolder, fSyncVersionFiles, DriveApp, SpreadsheetApp, fCreateNewCharacterSheet */
/* exported fCreateCharacter */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Character Management
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fRenameCharacter
   Purpose: The master orchestrator for the character renaming workflow.
   Assumptions: The Codex has a <Characters> sheet.
   Notes: Handles selection, prompting, and execution of the rename process.
   @returns {void}
*/
function fRenameCharacter() {
  const ssKey = 'Codex';
  const sheetName = 'Characters';
  fLoadSheetToArray(ssKey, sheetName);
  fBuildTagMaps(ssKey, sheetName);

  const destSS = SpreadsheetApp.getActiveSpreadsheet();
  const destSheet = destSS.getSheetByName(sheetName);
  const { arr, rowTags, colTags } = g[ssKey][sheetName];

  const startRow = rowTags.tablestart;
  const endRow = rowTags.tableend;
  const checkboxCol = colTags.checkbox;
  const charNameCol = colTags.charname;
  const csidCol = colTags.csid;
  const versionCol = colTags.version;

  // 1. Find the selected character (must be exactly one)
  const selectedCharacters = [];
  for (let r = startRow; r <= endRow; r++) {
    if (arr[r] && arr[r][checkboxCol] === true && arr[r][charNameCol]) {
      selectedCharacters.push({
        row: r + 1, // 1-based row
        name: arr[r][charNameCol],
        id: arr[r][csidCol],
        version: arr[r][versionCol],
      });
    }
  }

  // 2. Validate the selection
  if (selectedCharacters.length === 0) {
    fShowMessage('ℹ️ No Selection', 'Please check the box next to the character you wish to rename.');
    return;
  }
  if (selectedCharacters.length > 1) {
    fShowMessage('❌ Error', 'Multiple characters selected. Please select only one character to rename.');
    return;
  }

  const character = selectedCharacters[0];

  // 3. Get current names and prompt for a new one
  const file = DriveApp.getFileById(character.id);
  const currentFileName = file.getName();
  const currentSheetName = character.name;

  let promptMessage = `Current Name: ${currentSheetName}\n`;
  if (currentFileName !== currentSheetName) {
    promptMessage += `Current File Name: ${currentFileName}\n`;
  }
  promptMessage += '\nPlease enter the new name for this character:';

  const newBaseName = fPromptWithInput('Rename Character', promptMessage);

  if (!newBaseName) {
    fShowMessage('ℹ️ Canceled', 'Rename operation canceled.');
    return;
  }

  // 4. Process the new name (strip and re-apply correct version prefix)
  const cleanedName = newBaseName.replace(/^v\d+\s*/, '').trim();
  const finalName = `v${character.version} ${cleanedName}`;

  // 5. Execute the rename
  fShowToast(`Renaming to "${finalName}"...`, 'Rename Character');
  file.setName(finalName);

  const nameCell = destSheet.getRange(character.row, charNameCol + 1);
  const url = nameCell.getRichTextValue().getLinkUrl();
  const newLink = SpreadsheetApp.newRichTextValue().setText(finalName).setLinkUrl(url).build();
  nameCell.setRichTextValue(newLink);

  // 6. Final success message
  fEndToast();
  fShowMessage('✅ Success', `"${currentSheetName}" has been successfully renamed to "${finalName}".`);

} // End function fRenameCharacter

/* function fDeleteCharacter
   Purpose: The master orchestrator for the character deletion workflow.
   Assumptions: The Codex has a <Characters> sheet with a 'CheckBox' column.
   Notes: Handles single or multiple selections and provides a confirmation prompt before proceeding.
   @returns {void}
*/
function fDeleteCharacter() {
  const ssKey = 'Codex';
  const sheetName = 'Characters';
  fLoadSheetToArray(ssKey, sheetName);
  fBuildTagMaps(ssKey, sheetName);

  const destSS = SpreadsheetApp.getActiveSpreadsheet();
  const destSheet = destSS.getSheetByName(sheetName);
  const { arr, rowTags, colTags } = g[ssKey][sheetName];

  const startRow = rowTags.tablestart;
  const endRow = rowTags.tableend;
  const checkboxCol = colTags.checkbox;
  const charNameCol = colTags.charname;
  const csidCol = colTags.csid;

  if (checkboxCol === undefined || charNameCol === undefined || csidCol === undefined) {
    fShowMessage('❌ Error', 'The <Characters> sheet is missing a "CheckBox", "CharName", or "CSID" column tag.');
    return;
  }

  // 1. Find all checked characters
  const selectedCharacters = [];
  for (let r = startRow; r <= endRow; r++) {
    // Only consider rows that actually have a character name
    if (arr[r] && arr[r][checkboxCol] === true && arr[r][charNameCol]) {
      selectedCharacters.push({
        row: r + 1, // Store 1-based row for later
        name: arr[r][charNameCol],
        id: arr[r][csidCol],
      });
    }
  }

  // 2. Validate the selection and get user confirmation
  if (selectedCharacters.length === 0) {
    fShowMessage('ℹ️ No Selection', 'Please check the box next to the character(s) you wish to delete.');
    return;
  } else if (selectedCharacters.length === 1) {
    const charName = selectedCharacters[0].name;
    const promptMessage = `⚠️ Are you sure you wish to permanently DELETE the character "${charName}"?\n\nThis action cannot be undone.\n\nTo confirm, please type DELETE below.`;
    const confirmationText = fPromptWithInput('Confirm Deletion', promptMessage);
    if (confirmationText === null || confirmationText.toLowerCase().trim() !== 'delete') {
      fShowMessage('ℹ️ Canceled', 'Deletion has been canceled.');
      return;
    }
  } else {
    const names = selectedCharacters.map(c => `- ${c.name}`).join('\n');
    const promptMessage = `⚠️ Are you sure you wish to permanently DELETE the following ${selectedCharacters.length} characters?\n\n${names}\n\nThis action cannot be undone.\n\nTo confirm, please type DELETE ALL below.`;
    const confirmationText = fPromptWithInput('Confirm Bulk Deletion', promptMessage);
    if (confirmationText === null || confirmationText.toLowerCase().trim() !== 'delete all') {
      fShowMessage('ℹ️ Canceled', 'Deletion has been canceled.');
      return;
    }
  }

  // 3. Trash the Google Drive files
  selectedCharacters.forEach(character => {
    try {
      fShowToast(`🗑️ Trashing file for ${character.name}...`, 'Deleting');
      DriveApp.getFileById(character.id).setTrashed(true);
    } catch (e) {
      console.error(`Could not trash file with ID ${character.id} for character ${character.name}. It may have already been deleted. Error: ${e}`);
    }
  });

  // 4. Delete the spreadsheet rows using our new robust helper
  // Sort in reverse order to avoid index shifting issues
  selectedCharacters.sort((a, b) => b.row - a.row).forEach(character => {
    fDeleteTableRow(destSheet, character.row);
  });

  // 5. Final success message
  fEndToast();
  const deletedNames = selectedCharacters.map(c => c.name).join(', ');
  fShowMessage('✅ Success', `The following character(s) have been deleted:\n\n${deletedNames}`);

} // End function fDeleteCharacter

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
    fShowMessage('❌ Error', `Could not find the local master Character Sheet for Version ${version}. Please try syncing versions again.`);
    return;
  }

  // 2. Get the destination folder and copy the template
  fShowToast('Creating a new character sheet...', 'New Character');
  const charactersFolder = fGetOrCreateFolder('Characters', parentFolder);
  const csTemplateFile = DriveApp.getFileById(localCsId);
  const newCharSheet = csTemplateFile.makeCopy(charactersFolder);

  const characterName = fPromptWithInput('Name Your Character', 'Please enter a name for your new character:');

  if (!characterName) {
    newCharSheet.setTrashed(true);
    fEndToast();
    fShowMessage('ℹ️ Canceled', 'Character creation has been canceled.');
    return;
  }

  const versionedCharacterName = `v${version} ${characterName}`;
  newCharSheet.setName(versionedCharacterName);

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
  dataToWrite[colTags.charname - 1] = versionedCharacterName; // Use versioned name
  dataToWrite[colTags.rules - 1] = `v${version} Rules`;     // Use versioned rules text

  // Case 1: First character, table is empty.
  if (startRow === endRow && (!arr[startRow] || arr[startRow][charNameCol] === '')) {
    targetRow = startRow + 1;
    const targetRange = destSheet.getRange(targetRow, 2, 1, dataToWrite.length);
    targetRange.setValues([dataToWrite]);
  } else {
    // Case 2 & 3: One or more characters exist.
    targetRow = endRow + 2;
    destSheet.insertRowsAfter(endRow + 1, 1);

    const oldTagCell = destSheet.getRange(endRow + 1, 1);
    const oldTags = oldTagCell.getValue().toString().split(',').map(t => t.trim());
    const newTags = oldTags.filter(t => t.toLowerCase() !== 'tableend');
    oldTagCell.setValue(newTags.join(', '));
    destSheet.getRange(targetRow, 1).setValue('TableEnd');

    if (colTags.checkbox !== undefined) {
      const checkboxCol = colTags.checkbox + 1;
      const numRows = endRow - startRow + 1;
      destSheet.getRange(startRow + 1, checkboxCol, numRows, 1).uncheck();
    }
    const targetRange = destSheet.getRange(targetRow, 2, 1, dataToWrite.length);
    targetRange.setValues([dataToWrite]);
  }

  // 4. Format the new row appropriately
  if (colTags.checkbox !== undefined) {
    destSheet.getRange(targetRow, colTags.checkbox + 1).insertCheckboxes();
  }
  const link = SpreadsheetApp.newRichTextValue().setText(versionedCharacterName).setLinkUrl(newCharSheet.getUrl()).build();
  destSheet.getRange(targetRow, colTags.charname + 1).setRichTextValue(link);

  const rulesId = fGetSheetId(version, 'Rules');
  const rulesUrl = `https://docs.google.com/document/d/${rulesId}/`;
  const rulesLink = SpreadsheetApp.newRichTextValue().setText(`v${version} Rules`).setLinkUrl(rulesUrl).build();
  destSheet.getRange(targetRow, colTags.rules + 1).setRichTextValue(rulesLink);

  // 5. Final, corrected success message
  fEndToast();
  const successMessage = `✅ Success! Your new character, "${characterName}," has been created.\n\nA link has been added to your <Characters> sheet.`;
  fShowMessage('Character Created!', successMessage);
} // End function fCreateNewCharacterSheet


/* function fCreateLatestCharacter
   Purpose: Controller for creating a character using the latest available version without a prompt.
   Assumptions: None.
   Notes: Determines the latest version and calls the core character creation function. Triggers initial setup if needed.
   @returns {void}
*/
function fCreateLatestCharacter() {
  const ssKey = 'Codex';
  const sheetName = 'MyVersions';

  // 1. Check if the one-time setup is needed. This must be the first step.
  fLoadSheetToArray(ssKey, sheetName);
  fBuildTagMaps(ssKey, sheetName);
  let { arr, rowTags, colTags } = g[ssKey][sheetName];
  if (rowTags.tablestart === rowTags.tableend && (!arr[rowTags.tablestart] || arr[rowTags.tablestart][colTags.ssabbr] === '')) {
    fInitialSetup();
    // After setup, we MUST reload the sheet data to get the new information
    fLoadSheetToArray(ssKey, sheetName);
    fBuildTagMaps(ssKey, sheetName);
    let reloadedData = g[ssKey][sheetName];
    arr = reloadedData.arr;
    rowTags = reloadedData.rowTags;
    colTags = reloadedData.colTags;
  }

  // 2. Find the highest version number that has a CS file
  const versionsWithCS = arr.slice(rowTags.tablestart, rowTags.tableend + 1)
    .filter(row => row[colTags.ssabbr] === 'CS')
    .map(row => parseFloat(row[colTags.version]));

  if (versionsWithCS.length === 0) {
    fShowMessage('❌ Error', 'No versions with a Character Sheet (CS) were found in <MyVersions>.');
    return;
  }

  const latestVersion = Math.max(...versionsWithCS).toString();
  fCreateCharacterFromVersion(latestVersion);
} // End function fCreateLatestCharacter


/* function fCreateLegacyCharacter
   Purpose: Controller for creating a character from a list of older, non-latest versions.
   Assumptions: None.
   Notes: Prompts the user to select from a list of available legacy versions. Triggers initial setup if needed.
   @returns {void}
*/
function fCreateLegacyCharacter() {
  const ssKey = 'Codex';
  const sheetName = 'MyVersions';

  // 1. Check if the one-time setup is needed. This must be the first step.
  fLoadSheetToArray(ssKey, sheetName);
  fBuildTagMaps(ssKey, sheetName);
  let { arr, rowTags, colTags } = g[ssKey][sheetName];
  if (rowTags.tablestart === rowTags.tableend && (!arr[rowTags.tablestart] || arr[rowTags.tablestart][colTags.ssabbr] === '')) {
    fInitialSetup();
    // After setup, we MUST reload the sheet data to get the new information
    fLoadSheetToArray(ssKey, sheetName);
    fBuildTagMaps(ssKey, sheetName);
    let reloadedData = g[ssKey][sheetName];
    arr = reloadedData.arr;
    rowTags = reloadedData.rowTags;
    colTags = reloadedData.colTags;
  }

  // 2. Find and prompt for legacy versions
  const versionsWithCS = arr.slice(rowTags.tablestart, rowTags.tableend + 1)
    .filter(row => row[colTags.ssabbr] === 'CS')
    .map(row => parseFloat(row[colTags.version]));

  if (versionsWithCS.length === 0) {
    fShowMessage('❌ Error', 'No versions with a Character Sheet (CS) were found in <MyVersions>.');
    return;
  }

  const latestVersion = Math.max(...versionsWithCS);
  const legacyVersions = [...new Set(versionsWithCS.filter(v => v < latestVersion).map(String))];

  if (legacyVersions.length === 0) {
    fShowMessage('ℹ️ No Legacy Versions', 'No older legacy versions are available to choose from.');
    return;
  }

  const promptMessage = `Please enter the legacy game version you would like to use.\n\nAvailable versions:\n${legacyVersions.join(', ')}`;
  const selectedVersion = fPromptWithInput('Select Legacy Version', promptMessage);

  if (selectedVersion === null) {
    fShowMessage('ℹ️ Canceled', 'Character creation has been canceled.');
    return;
  }

  if (!legacyVersions.includes(selectedVersion)) {
    fShowMessage('❌ Error', `Invalid version selected. Please enter one of the available versions: ${legacyVersions.join(', ')}`);
    return;
  }

  fCreateCharacterFromVersion(selectedVersion);
} // End function fCreateLegacyCharacter


/* function fCreateCharacterFromVersion
   Purpose: The core logic for character creation, now triggered by a specific version.
   Assumptions: The initial setup has already been run and a valid version is provided.
   Notes: This is the generic helper function called by the menu controllers.
   @param {string} selectedVersion - The version of the character to create.
   @returns {void}
*/
function fCreateCharacterFromVersion(selectedVersion) {
  if (!selectedVersion) {
    fShowMessage('❌ Error', 'No version was provided for character creation.');
    return;
  }

  // Create the new character sheet and log it.
  const parentFolder = fGetOrCreateFolder('MetaScape Flex');
  fCreateNewCharacterSheet(selectedVersion, parentFolder);

} // End function fCreateCharacterFromVersion


