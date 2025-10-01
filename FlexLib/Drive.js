/* global DriveApp, PropertiesService */
/* exported fGetOrCreateFolder, fSyncVersionFiles */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Google Drive Management
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


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
  const masterCopiesFolder = fGetOrCreateFolder('Master Copies - DO NOT DELETE', parentFolder);
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

      const fileName = `MASTER_${ssAbbr} - DO NOT DELETE`;
      fShowToast(`⏳ Copying ${ssAbbr} file...`, 'Syncing Files');
      const newFile = DriveApp.getFileById(masterId).makeCopy(fileName, masterCopiesFolder);
      localCache[version][ssAbbr] = newFile.getId();
      fShowToast(`✅ Copied ${ssAbbr} successfully!`, 'Syncing Files');
    }
  });

  properties.setProperty('localFileCache', JSON.stringify(localCache));
} // End function fSyncVersionFiles