/* global fShowMessage, fBuildTagMaps, g, fPromptWithInput, fGetSheetId, fGetOrCreateFolder, fSyncVersionFiles, DriveApp, SpreadsheetApp, fCreateNewCharacterSheet */
/* exported fCreateCharacter */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Character Management
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fCreateNewCharacterSheet
   Purpose: Creates and names a new character sheet from the local master and logs it in the Codex.
   Assumptions: The required master files for the selected version have already been synced.
   Notes: This is the final step in the character creation workflow.
   @param {string} version - The game version for the new character (e.g., '3').
   @param {GoogleAppsScript.Drive.Folder} parentFolder - The "MetaScape Flex" folder object.
   @returns {void}
*/
function fCreateNewCharacterSheet(version, parentFolder) {
  // 1. Get the local CS template ID from PropertiesService
  const properties = PropertiesService.getScriptProperties();
  const localCache = JSON.parse(properties.getProperty('localFileCache') || '{}');
  const localCsId = localCache[version] ? localCache[version]['CS'] : null;

  if (!localCsId) {
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
  // Note: This data array does NOT include the tag column.
  const rulesId = g.sheetIDs[version]['Rules'].ssid;
  const rulesUrl = `https://docs.google.com/spreadsheets/d/${rulesId}/`;
  const dataToWrite = [];
  dataToWrite[colTags.csid - 1] = newCharSheet.getId();
  dataToWrite[colTags.version - 1] = version;
  dataToWrite[colTags.checkbox - 1] = true;
  dataToWrite[colTags.charname - 1] = characterName;
  dataToWrite[colTags.rules - 1] = rulesUrl;


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

  // 5. Final, corrected success message
  const successMessage = `✅ Success! Your new character, "${characterName}," has been created.\n\nA link has been added to your <Characters> sheet.`;
  fShowMessage('Character Created!', successMessage);
} // End function fCreateNewCharacterSheet


/* function fCreateCharacter
   Purpose: The master orchestrator for the entire character creation workflow.
   Assumptions: The Codex has a <Versions> sheet with a 'version' column tag.
   Notes: This phase prompts the user, syncs files, and creates the new character sheet.
   @returns {void}
*/
function fCreateCharacter() {
  // 1. Load available versions
  const ssKey = 'Codex';
  const sheetName = 'Versions';
  fBuildTagMaps(ssKey, sheetName);

  const { arr, rowTags, colTags } = g[ssKey][sheetName];
  const startRow = rowTags.tablestart;
  const endRow = rowTags.tableend;
  const versionCol = colTags.version;

  const availableVersions = [...new Set(arr.slice(startRow, endRow + 1).map(row => String(row[versionCol])))];

  if (availableVersions.length === 0) {
    fShowMessage('❌ Error', 'No game versions were found in the <Versions> sheet.');
    return;
  }

  // 2. Prompt user for selection
  const promptMessage = `Please enter the game version you would like to use.\n\nAvailable versions:\n${availableVersions.join(', ')}`;
  const selectedVersion = fPromptWithInput('Select Game Version', promptMessage);

  // 3. Handle response
  if (selectedVersion === null) {
    fShowMessage('ℹ️ Canceled', 'Character creation has been canceled.');
    return;
  }

  if (!availableVersions.includes(selectedVersion)) {
    fShowMessage('❌ Error', `Invalid version selected. Please enter one of the available versions: ${availableVersions.join(', ')}`);
    return;
  }

  // 4. Ensure master files are synced for the selected version.
  fShowMessage('Character Creation', `⏳ Syncing master files for Version ${selectedVersion}...`);
  // First, ensure the ID cache is loaded by calling the getter. This is the critical first step.
  fGetSheetId(selectedVersion, 'CS');
  // Now, explicitly pass the loaded data to the sync function.
  const parentFolder = fGetOrCreateFolder('MetaScape Flex');
  fSyncVersionFiles(selectedVersion, parentFolder, g.sheetIDs[selectedVersion]);
  fShowMessage('Character Creation', `✅ File sync for Version ${selectedVersion} complete.`);

  // 5. Create the new character sheet and log it.
  fCreateNewCharacterSheet(selectedVersion, parentFolder);
} // End function fCreateCharacter