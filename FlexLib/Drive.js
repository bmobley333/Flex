/* global DriveApp, PropertiesService */
/* exported fGetOrCreateFolder, fSyncVersionFiles, fGetSubFolder */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Google Drive Management
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fGetSubFolder
   Purpose: A robust, self-healing "health check" to get a required subfolder using its stored ID.
   Assumptions: The FlexFolderID and the specific subfolder's ID are stored in the Codex's <Data> sheet.
   Notes: This is the definitive gatekeeper for all subfolder access. If a folder is missing, it recreates it.
   @param {string} folderTag - The row tag for the folder ID in the <Data> sheet (e.g., 'characterfolderid').
   @param {string} defaultFolderName - The user-facing name to give the folder if it needs to be recreated (e.g., 'Characters').
   @returns {GoogleAppsScript.Drive.Folder|null} The subfolder object, or null if a critical error occurs.
*/
function fGetSubFolder(folderTag, defaultFolderName) {
  const codexSS = fGetCodexSpreadsheet();
  const dataSheet = codexSS.getSheetByName('Data');
  const { arr, rowTags, colTags } = fGetSheetData('Codex', 'Data', codexSS, true); // Force refresh
  const dataCol = colTags.data;

  // 1. Get the Main Flex Folder first, as it's the parent for everything.
  const flexFolderIdRow = rowTags.flexfolderid;
  if (flexFolderIdRow === undefined) {
    fShowMessage('❌ Error', 'Could not find the `FlexFolderID` tag in your <Data> sheet. Please run the setup again.');
    return null;
  }
  const flexFolderId = arr[flexFolderIdRow][dataCol];
  let mainFolder;
  try {
    mainFolder = DriveApp.getFolderById(flexFolderId);
  } catch (e) {
    fShowMessage('❌ Error', 'Could not access the main "💪 My Flex" folder. It may have been deleted. Please run the setup again to restore it.');
    return null;
  }

  // 2. Get the specific subfolder's ID.
  const subFolderIdRow = rowTags[folderTag];
  if (subFolderIdRow === undefined) {
    fShowMessage('❌ Error', `Could not find the folder tag "${folderTag}" in your <Data> sheet. Your Codex may be out of date.`);
    return null;
  }
  const subFolderId = arr[subFolderIdRow][dataCol];

  // 3. Try to access the folder by its ID.
  if (subFolderId) {
    try {
      // --- Happy Path ---
      return DriveApp.getFolderById(subFolderId);
    } catch (e) {
      // --- Self-Heal Path ---
      // The ID is logged but the folder was deleted. Recreate it.
      fShowToast(`⏳ The "${defaultFolderName}" folder was missing. Restoring it for you...`, 'System Health');
      const newFolder = mainFolder.createFolder(defaultFolderName);
      const newFolderId = newFolder.getId();
      dataSheet.getRange(subFolderIdRow + 1, dataCol + 1).setValue(newFolderId);
      return newFolder;
    }
  } else {
    // --- Self-Heal Path (First Run) ---
    // The ID was never logged. Create the folder for the first time.
    fShowToast(`⏳ Creating the "${defaultFolderName}" folder for you...`, 'System Health');
    const newFolder = mainFolder.createFolder(defaultFolderName);
    const newFolderId = newFolder.getId();
    dataSheet.getRange(subFolderIdRow + 1, dataCol + 1).setValue(newFolderId);
    return newFolder;
  }
} // End function fGetSubFolder

/* function fMoveFileToFolder
   Purpose: Moves a file to a specified folder if it's not already there.
   Assumptions: The user has granted DriveApp permissions.
   Notes: This helps organize the user's Drive.
   @param {GoogleAppsScript.Drive.File} file - The file object to move.
   @param {GoogleAppsScript.Drive.Folder} folder - The destination folder object.
   @returns {void}
*/
function fMoveFileToFolder(file, folder) {
  const parents = file.getParents();
  let isAlreadyInFolder = false;
  if (parents.hasNext()) {
    const parent = parents.next();
    if (parent.getId() === folder.getId()) {
      isAlreadyInFolder = true;
    }
  }

  if (!isAlreadyInFolder) {
    file.moveTo(folder);
  }
} // End function fMoveFileToFolder



/* function fGetOrCreateFolder
   Purpose: Finds a folder by name in a given location, or creates it if it doesn't exist.
   Assumptions: The user has granted the necessary DriveApp permissions.
   Notes: If parentFolder is null, it searches/creates in the root of the user's Drive.
   @param {string} folderName - The name of the folder to find or create.
   @param {GoogleAppsScript.Drive.Folder} [parentFolder=null] - The folder to search within. Defaults to root.
   @returns {GoogleAppsScript.Drive.Folder} The folder object.
*/
function fGetOrCreateFolder(folderName, parentFolder = null) {
  const root = parentFolder || DriveApp.getRootFolder();
  const folders = root.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  } else {
    return root.createFolder(folderName);
  }
} // End function fGetOrCreateFolder

/* function fSyncVersionFiles
   Purpose: Copies specific master files for a given version to the user's local Drive.
   Assumptions: The filesToSync object is correctly passed as a parameter.
   Notes: Uses PropertiesService to ensure files are only copied once. Skips Ver and Codex.
   @param {string} version - The version number to sync files for (e.g., '3').
   @param {GoogleAppsScript.Drive.Folder} parentFolder - The main "MetaScape Flex" folder.
   @param {object} filesToSync - The object containing the file info for the specified version.
   @returns {void}
*/
function fSyncVersionFiles(version, parentFolder, filesToSync) {
  const masterCopiesFolder = fGetOrCreateFolder('Master Copies', parentFolder);
  const properties = PropertiesService.getScriptProperties();
  const localCache = JSON.parse(properties.getProperty('localFileCache') || '{}');

  if (!localCache[version]) {
    localCache[version] = {};
  }

  // Define which files are necessary for a player's local setup
  const requiredFiles = ['CS', 'DB', 'Rules'];

  requiredFiles.forEach(ssAbbr => {
    // Check if the file for this version has been copied already AND that we have info for it
    if (!localCache[version][ssAbbr] && filesToSync[ssAbbr]) {
      const masterId = filesToSync[ssAbbr].ssid;
      if (!masterId) return; // Skip if the ID doesn't exist for some reason

      const fileName = `MASTER_${ssAbbr}`;
      fShowToast(`⏳ Copying ${ssAbbr} file...`, 'Syncing Files');
      const newFile = DriveApp.getFileById(masterId).makeCopy(fileName, masterCopiesFolder);
      localCache[version][ssAbbr] = newFile.getId();
      fShowToast(`✅ Copied ${ssAbbr} successfully!`, 'Syncing Files');
    }
  });

  properties.setProperty('localFileCache', JSON.stringify(localCache));
} // End function fSyncVersionFiles