/* global fShowMessage */
/* exported fCreateCharacter */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Character Management
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fCreateCharacter
   Purpose: The master orchestrator for the entire character creation workflow.
   Assumptions: The Codex has a <Versions> sheet with a 'version' column tag.
   Notes: This phase prompts the user, syncs files, and moves/renames the Codex.
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

  // 5. Move and rename this Codex file
  const thisFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  fMoveFileToFolder(thisFile, parentFolder);
  thisFile.setName("Player's Codex");

  fShowMessage('Character Creation', `✅ File sync for Version ${selectedVersion} complete.`);
} // End function fCreateCharacter