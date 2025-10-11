/* global fShowMessage, DriveApp, SpreadsheetApp, g, fNormalizeTags, fLoadSheetToArray, fBuildTagMaps, MimeType, fEmbedCodexId */
/* exported fLogLocalFileCopy */

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

  // 1. Create Folder Structure & Store Folder IDs
  fShowToast('Creating Google Drive folders...', '⚙️ Setup');
  const parentFolder = fGetOrCreateFolder('💪 My Flex');
  const masterCopiesFolder = fGetOrCreateFolder('Master Copies', parentFolder);
  const charactersFolder = fGetOrCreateFolder('Characters', parentFolder);
  const customAbilitiesFolder = fGetOrCreateFolder('Custom Abilities', parentFolder);

  const codexSS = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = codexSS.getSheetByName('Data');
  if (dataSheet) {
    const { rowTags, colTags } = fGetSheetData('Codex', 'Data', codexSS, true);
    const dataCol = colTags.data;

    const folderIdMap = {
      flexfolderid: parentFolder.getId(),
      mastercopiesfolderid: masterCopiesFolder.getId(),
      characterfolderid: charactersFolder.getId(),
      custabilfolderid: customAbilitiesFolder.getId(),
    };

    for (const rowTag in folderIdMap) {
      const rowIndex = rowTags[rowTag];
      if (rowIndex !== undefined && dataCol !== undefined) {
        dataSheet.getRange(rowIndex + 1, dataCol + 1).setValue(folderIdMap[rowTag]);
      }
    }
  }

  // 2. Move and Rename this Codex
  fShowToast('Organizing your Codex file...', '⚙️ Setup');
  const thisFile = DriveApp.getFileById(codexSS.getId());
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
  fSyncAllVersionFiles(sourceData, masterCopiesFolder);

  // 5. The setup is now complete. Custom abilities are created on-demand by the user.
  fEndToast();
  const successMessage = 'Your Player\'s Codex is now ready to use.\n\nPlease bookmark this Player\'s Codex file (and you can also find it in your Google Drive under the "💪 My Flex" folder).';
  fShowMessage('✅ Setup Complete!', successMessage);
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

  // --- NEW LOGIC ---
  // Use the 'Header' tag to find the start of the data.
  const headerRow = sourceRowTags.header;
  if (headerRow === undefined) {
    throw new Error('Could not find "Header" row tag in the master <Versions> sheet.');
  }
  const startRow = headerRow + 1;

  // 2. Define the columns we need to extract from the source sheet
  const versionCol = sourceColTags.version;
  const releaseDateCol = sourceColTags.releasedate;
  const playerNeedsCol = sourceColTags.playerneeds;
  const fullNameCol = sourceColTags.ssfullname;
  const abbrCol = sourceColTags.ssabbr;
  const idCol = sourceColTags.ssid;

  // 3. Loop through each row of the source data table and process it
  for (let r = startRow; r < sourceData.length; r++) {
    const rowData = sourceData[r];

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
    const fileName = `v${version} MASTER_${ssAbbr}`;
    const newFile = DriveApp.getFileById(masterId).makeCopy(fileName, masterCopiesFolder);

    // Only try to embed the Codex ID if the new file is a spreadsheet.
    if (newFile.getMimeType() === MimeType.GOOGLE_SHEETS) {
      const newSS = SpreadsheetApp.openById(newFile.getId());
      fEmbedCodexId(newSS);
    }

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
   Notes: Uses a "Header"-based approach and propagates formatting.
   @param {object} logData - An object containing the data for the new file.
   @returns {void}
*/
function fLogLocalFileCopy(logData) {
  const ssKey = 'Codex';
  const sheetName = 'MyVersions';
  const destSS = SpreadsheetApp.getActiveSpreadsheet();
  const destSheet = destSS.getSheetByName(sheetName);

  const { arr, colTags, rowTags } = fGetSheetData(ssKey, sheetName, destSS, true);
  const headerRow = rowTags.header;
  const lastRow = destSheet.getLastRow();
  let targetRow;

  if (headerRow === undefined) {
    throw new Error(`Could not find a "Header" tag in the <${sheetName}> sheet.`);
  }

  const dataToWrite = [];
  dataToWrite[colTags.version - 1] = logData.version;
  dataToWrite[colTags.releasedate - 1] = logData.releaseDate;
  dataToWrite[colTags.ssfullname - 1] = logData.ssFullName;
  dataToWrite[colTags.ssabbr - 1] = logData.ssAbbr;
  dataToWrite[colTags.ssid - 1] = logData.ssID;

  const firstDataRowIndex = headerRow + 1;
  const templateRow = firstDataRowIndex + 1; // 1-based template row number
  const ssAbbrCol = colTags.ssabbr;

  // Case 1: The table is empty. Write data to the pre-formatted first row.
  if (arr.length <= firstDataRowIndex || !arr[firstDataRowIndex][ssAbbrCol]) {
    targetRow = templateRow;
  } else {
    // Case 2: The table has data. Insert a new row and copy formatting.
    targetRow = lastRow + 1;
    destSheet.insertRowsAfter(lastRow, 1);

    const formatSourceRange = destSheet.getRange(templateRow, 1, 1, destSheet.getMaxColumns());
    const formatDestRange = destSheet.getRange(targetRow, 1, 1, destSheet.getMaxColumns());
    formatSourceRange.copyTo(formatDestRange, { formatOnly: true });
  }

  // Write the data starting from the second column to preserve row tags
  const targetRange = destSheet.getRange(targetRow, 2, 1, dataToWrite.length);
  targetRange.setValues([dataToWrite]);
} // End function fLogLocalFileCopy