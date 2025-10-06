/* global fShowMessage, DriveApp, SpreadsheetApp, g, fNormalizeTags, fLoadSheetToArray, fBuildTagMaps */
/* exported fInitialSetup */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - First-Time User Setup
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fInitialSetup
   Purpose: The master orchestrator for the entire one-time, first-use setup process for a new player.
   Assumptions: This is being run from a fresh copy of the Codex template.
   Notes: This function creates folders, moves the Codex, and triggers the sync of all master files.
   @returns {void}
*/
function fInitialSetup() {
  const welcomeMessage = 'Welcome to Flex! This will perform a one-time setup to prepare your Player\'s Codex.\n\n⚠️ This process may take several minutes to complete. Please do not close this spreadsheet or navigate away until you see the "Setup Complete!" message.';
  fShowMessage('👋 Welcome!', welcomeMessage);

  // 1. Create Folder Structure
  fShowToast('Creating Google Drive folders...', '⚙️ Setup');
  const parentFolder = fGetOrCreateFolder('MetaScape Flex');
  fGetOrCreateFolder('Master Copies - DO NOT DELETE', parentFolder);
  fGetOrCreateFolder('Characters', parentFolder); // Create the new Characters sub-folder

  // 2. Move and Rename this Codex
  fShowToast('Organizing your Codex file...', '⚙️ Setup');
  const thisFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  fMoveFileToFolder(thisFile, parentFolder);
  thisFile.setName("Player's Codex");

  // 3. Get Master Version Data
  fShowToast('Fetching the latest version list...', '⚙️ Setup');
  const sourceSS = SpreadsheetApp.openById(g.MASTER_VER_ID);
  const sourceSheet = sourceSS.getSheetByName('Versions');
  if (!sourceSheet) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find the master <Versions> sheet. Please contact the administrator.');
    return;
  }
  const sourceData = sourceSheet.getDataRange().getValues();

  // 4. Sync all files and log them to the local <MyVersions> sheet
  const masterCopiesFolder = fGetOrCreateFolder('Master Copies - DO NOT DELETE', parentFolder);
  fSyncAllVersionFiles(sourceData, masterCopiesFolder);

  fEndToast();
  fShowMessage('✅ Setup Complete!', 'Your Player\'s Codex is now ready to use.');
} // End function fInitialSetup


/* function fSyncAllVersionFiles
   Purpose: Reads master version data, copies only files marked as PlayerNeeds, and logs them.
   Assumptions: The sourceData is a 2D array from the master Ver sheet.
   Notes: This is the core file synchronization engine for the initial setup.
   @param {Array<Array<string>>} sourceData - The full data from the master "Ver" sheet.
   @param {GoogleAppsScript.Drive.Folder} masterCopiesFolder - The "Master Copies" folder object.
   @returns {void}
*/
function fSyncAllVersionFiles(sourceData, masterCopiesFolder) {
  // 1. Build temporary tag maps to understand the source data structure
  const sourceRowTags = {};
  const sourceColTags = {};
  sourceData[0].forEach((tag, c) => fNormalizeTags(tag).forEach(t => (sourceColTags[t] = c)));
  sourceData.forEach((row, r) => fNormalizeTags(row[0]).forEach(t => (sourceRowTags[t] = r)));

  const startRow = sourceRowTags.tablestart;
  const endRow = sourceRowTags.tableend;

  // 2. Define the columns we need to extract from the source sheet
  const versionCol = sourceColTags.version;
  const releaseDateCol = sourceColTags.releasedate;
  const playerNeedsCol = sourceColTags.playerneeds; // Updated from ismaster
  const fullNameCol = sourceColTags.ssfullname;
  const abbrCol = sourceColTags.ssabbr;
  const idCol = sourceColTags.ssid;

  // 3. Loop through each row of the source data table and process it
  for (let r = startRow; r <= endRow; r++) {
    const rowData = sourceData[r];

    // --- NEW LOGIC ---
    // Only copy the file if the 'PlayerNeeds' column is TRUE
    if (rowData[playerNeedsCol] !== true) {
      continue; // Skip this file
    }

    const masterId = rowData[idCol];
    const ssAbbr = rowData[abbrCol];
    const version = rowData[versionCol];
    if (!masterId || !ssAbbr) continue;

    fShowToast(`⏳ Copying ${ssAbbr} (Version ${version})...`, '⚙️ Setup');

    // 4. Make the copy with the new versioned file name
    const fileName = `v${version} MASTER_${ssAbbr} - DO NOT DELETE`;
    const newFile = DriveApp.getFileById(masterId).makeCopy(fileName, masterCopiesFolder);

    // 5. Prepare the data object to be logged
    const logData = {
      version: version,
      releaseDate: rowData[releaseDateCol],
      ssFullName: rowData[fullNameCol],
      ssAbbr: ssAbbr,
      ssID: newFile.getId(), // Log the NEW file's ID
    };

    // 6. Log the new file's info into the <MyVersions> sheet
    fLogLocalFileCopy(logData);
    fShowToast(`✅ Copied ${ssAbbr} (Version ${version}) successfully!`, '⚙️ Setup');
  }
} // End function fSyncAllVersionFiles


/* function fLogLocalFileCopy
   Purpose: Writes the details of a newly created local master file into the player's <MyVersions> sheet.
   Assumptions: The logData object contains all necessary keys.
   Notes: Contains the robust TableStart/TableEnd logic for a growing table.
   @param {object} logData - An object containing the data for the new file.
   @returns {void}
*/
function fLogLocalFileCopy(logData) {
  const ssKey = 'Codex';
  const sheetName = 'MyVersions';
  // Ensure the latest sheet data is loaded before we operate on it
  fLoadSheetToArray(ssKey, sheetName);
  fBuildTagMaps(ssKey, sheetName);

  const destSS = SpreadsheetApp.getActiveSpreadsheet();
  const destSheet = destSS.getSheetByName(sheetName);
  const { arr, rowTags, colTags } = g[ssKey][sheetName];

  const startRow = rowTags.tablestart;
  const endRow = rowTags.tableend;
  const ssAbbrCol = colTags.ssabbr; // Use a column to check if the first row is empty

  let targetRow;

  // Prepare the data to be written, matching the column order
  const dataToWrite = [];
  dataToWrite[colTags.version - 1] = logData.version;
  dataToWrite[colTags.releasedate - 1] = logData.releaseDate;
  // 'ismaster' column is removed
  dataToWrite[colTags.ssfullname - 1] = logData.ssFullName;
  dataToWrite[colTags.ssabbr - 1] = logData.ssAbbr;
  dataToWrite[colTags.ssid - 1] = logData.ssID;


  // Case 1: First file, table is empty.
  if (startRow === endRow && (!arr[startRow] || arr[startRow][ssAbbrCol] === '')) {
    targetRow = startRow + 1;
    // Data is written starting from the second column to preserve tags
    const targetRange = destSheet.getRange(targetRow, 2, 1, dataToWrite.length);
    targetRange.setValues([dataToWrite]);
  } else {
    // Case 2 & 3: One or more files already logged.
    targetRow = endRow + 2;
    destSheet.insertRowsAfter(endRow + 1, 1);

    // Move the 'TableEnd' tag
    const oldTagCell = destSheet.getRange(endRow + 1, 1);
    const oldTags = oldTagCell.getValue().toString().split(',').map(t => t.trim());
    const newTags = oldTags.filter(t => t.toLowerCase() !== 'tableend');
    oldTagCell.setValue(newTags.join(', '));
    destSheet.getRange(targetRow, 1).setValue('TableEnd');

    // Write the data starting from the second column
    const targetRange = destSheet.getRange(targetRow, 2, 1, dataToWrite.length);
    targetRange.setValues([dataToWrite]);
  }
} // End function fLogLocalFileCopy