/* global fShowMessage */
/* exported fCreateCharacter */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Character Management
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fCreateCharacter
   Purpose: The master orchestrator for the entire character creation workflow.
   Assumptions: The Codex has a <Versions> sheet with a 'version' column tag.
   Notes: This phase prompts the user to select a version.
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

  // Use a Set to get only unique version numbers, converting them to Strings for reliable comparison.
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

  fShowMessage('Character Creation', `✅ You selected Version ${selectedVersion}.`);
} // End function fCreateCharacter