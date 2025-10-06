/* global g, fGetSheetId, SpreadsheetApp, fBuildTagMaps, fShowMessage, fShowToast */
/* exported fBuildPowers */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Power List Generation
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fUpdatePowerTablesList
   Purpose: Updates the <Choose Powers> sheet with a unique list of all TableNames from the master DB <Powers> sheet.
   Assumptions: The user is running this from a Character Sheet that contains a sheet named <Choose Powers>.
   Notes: This is a designer-facing function for template maintenance. It dynamically adds/removes rows as needed.
   @returns {void}
*/
function fUpdatePowerTablesList() {
  // --- SECURITY CHECK ---
  if (Session.getActiveUser().getEmail() !== g.ADMIN_EMAIL) {
    fShowMessage('❌ Access Denied', 'This function is for designer use only.');
    return;
  }
  // --- END SECURITY CHECK ---

  fShowToast('⏳ Updating power table list...', 'Sync Power Tables');

  // 1. Get the ID for the master DB spreadsheet
  const dbId = fGetMasterSheetId(g.CURRENT_VERSION, 'DB');
  if (!dbId) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find the master "DB" spreadsheet ID in the master <Versions> sheet.');
    return;
  }
  const sourceSS = SpreadsheetApp.openById(dbId);

  const { arr, rowTags, colTags } = fGetSheetData('DB', 'Powers', sourceSS);

  // 2. Open the source DB <Powers> sheet and get all TableName values
  const sourceSheet = sourceSS.getSheetByName('Powers');
  if (!sourceSheet) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find the <Powers> sheet in the master DB file.');
    return;
  }

  const startRow = rowTags.header + 1; // Data starts after the header
  const tableNameCol = colTags.tablename;

  const allTableNames = arr.slice(startRow).map(row => row[tableNameCol]);
  const uniqueTableNames = [...new Set(allTableNames)].sort();

  // 3. Get the destination <Choose Powers> sheet and its properties
  const destSS = SpreadsheetApp.getActiveSpreadsheet();
  const destSheet = destSS.getSheetByName('Choose Powers');
  if (!destSheet) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find the <Choose Powers> sheet in this spreadsheet.');
    return;
  }

  const { rowTags: destRowTags, colTags: destColTags } = fGetSheetData('CS', 'Choose Powers', destSS, true);

  const destHeaderRow = destRowTags.header;
  const destTableNameCol = destColTags.tablename + 1;
  const destCheckboxCol = destColTags.isactive + 1;

  // 4. Adjust rows to match the new list size
  const lastRow = destSheet.getLastRow();
  const currentRowCount = lastRow - destHeaderRow;
  const newRowCount = uniqueTableNames.length;

  if (newRowCount > currentRowCount) {
    destSheet.insertRowsAfter(lastRow, newRowCount - currentRowCount);
  } else if (newRowCount < currentRowCount) {
    destSheet.deleteRows(destHeaderRow + 1 + newRowCount, currentRowCount - newRowCount);
  }

  // 5. Write the new data and insert checkboxes
  if (newRowCount > 0) {
    const tableNameData = uniqueTableNames.map(name => [name]); // Convert to 2D array
    destSheet.getRange(destHeaderRow + 1, destTableNameCol, newRowCount, 1).setValues(tableNameData);
    destSheet.getRange(destHeaderRow + 1, destCheckboxCol, newRowCount, 1).insertCheckboxes();
  }

  fEndToast();
  fShowMessage('✅ Success', `The <Choose Powers> sheet has been updated with ${newRowCount} power tables.`);
} // End function fUpdatePowerTablesList


/* function fFilterPowers
   Purpose: Builds custom power selection dropdowns on the Character Sheet based on the player's choices in <Choose Powers>.
   Assumptions: The user is running this from a Character Sheet. The CS has a <Choose Powers> and a game sheet with power dropdown tags.
   Notes: This is the primary player-facing function for customizing their power list. It now also populates a local cache sheet.
   @returns {void}
*/
function fFilterPowers() {
  fShowToast('⏳ Filtering power lists...', 'Filter Powers');

  const csSS = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Read the player's choices from the <Choose Powers> sheet, forcing a refresh.
  const { arr: choicesArr, rowTags: choicesRowTags, colTags: choicesColTags } = fGetSheetData('CS', 'Choose Powers', csSS, true);

  const choicesSheet = csSS.getSheetByName('Choose Powers');
  if (!choicesSheet) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find the <Choose Powers> sheet.');
    return;
  }

  const choicesHeaderRow = choicesRowTags.header;
  const tableNameCol = choicesColTags.tablename;
  const isActiveCol = choicesColTags.isactive;

  const selectedTables = choicesArr
    .slice(choicesHeaderRow + 1)
    .filter(row => row[isActiveCol] === true)
    .map(row => row[tableNameCol]);

  if (selectedTables.length === 0) {
    fEndToast();
    fShowMessage('ℹ️ No Filters Selected', 'Please check one or more boxes on the <Choose Powers> sheet before filtering.');
    return;
  }

  // 2. Fetch all powers from the player's local DB copy.
  const dbId = fGetSheetId(g.CURRENT_VERSION, 'DB');
  if (!dbId) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find the ID for the "DB" spreadsheet in <MyVersions>.');
    return;
  }
  const dbSS = SpreadsheetApp.openById(dbId);

  // --- REFACTORED ---
  const { arr: allPowers, rowTags: dbRowTags, colTags: dbColTags } = fGetSheetData('DB', 'Powers', dbSS);
  // --- END REFACTORED ---

  // 3. Filter the powers in-memory.
  const filteredPowers = allPowers
    .slice(dbRowTags.header + 1)
    .filter(row => selectedTables.includes(row[dbColTags.tablename]));

  if (filteredPowers.length === 0) {
    fEndToast();
    fShowMessage('⚠️ No Powers Found', 'No powers matched your selected filters. The dropdowns will be empty.');
  }

  // 4. Populate the <PowerDataCache> sheet
  const cacheSheet = csSS.getSheetByName('PowerDataCache');
  if (!cacheSheet) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find the <PowerDataCache> sheet.');
    return;
  }
  cacheSheet.clear(); // Clear old data
  const dbHeader = allPowers[dbRowTags.header];
  const dataToCache = [dbHeader, ...filteredPowers]; // Include header for easy lookups later
  cacheSheet.getRange(1, 1, dataToCache.length, dataToCache[0].length).setValues(dataToCache);
  fShowToast('⚡ Power data cached locally.', 'Filter Powers');


  // 5. Build and apply the validation rule to the <Game> sheet.
  const filteredPowerList = filteredPowers.map(row => row[dbColTags.dropdown]);
  const gameSheet = csSS.getSheetByName('Game');
  if (!gameSheet) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find the <Game> sheet.');
    return;
  }

  // --- REFACTORED ---
  const { rowTags: gameRowTags, colTags: gameColTags } = fGetSheetData('CS', 'Game', csSS);
  // --- END REFACTORED ---

  const startRow = gameRowTags.powertablestart + 1;
  const endRow = gameRowTags.powertableend + 1;
  const numRows = endRow - startRow + 1;

  const rule = SpreadsheetApp.newDataValidation().requireValueInList(filteredPowerList, true).setAllowInvalid(false).build();
  const dropDownCols = Object.keys(gameColTags).filter(tag => tag.startsWith('powerdropdown'));
  dropDownCols.forEach(tag => {
    const colIndex = gameColTags[tag] + 1;
    gameSheet.getRange(startRow, colIndex, numRows, 1).setDataValidation(rule);
  });

  fEndToast();
  fShowMessage('✅ Success!', `Your power selection dropdowns have been updated with ${filteredPowerList.length} powers.`);
} // End function fFilterPowers

/* function fBuildPowers
   Purpose: The master function to rebuild the <Powers> sheet in the DB file from the master Tables file.
   Assumptions: The user is running this from the DB spreadsheet.
   Notes: This is a destructive and regenerative process that now reads from multiple source sheets.
   @returns {void}
*/
function fBuildPowers() {
  // --- SECURITY CHECK ---
  if (Session.getActiveUser().getEmail() !== g.ADMIN_EMAIL) {
    fShowMessage('❌ Access Denied', 'This function is for designer use only.');
    return;
  }
  // --- END SECURITY CHECK ---

  fShowToast('⏳ Initializing power build...', 'Build Powers');

  // 1. Get the ID of the master Tables spreadsheet from the master Ver sheet
  const tablesId = fGetMasterSheetId(g.CURRENT_VERSION, 'Tbls');
  if (!tablesId) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find the ID for the "Tbls" spreadsheet in the master <Versions> sheet.');
    return;
  }

  // 2. Define source and destination details
  const sourceSS = SpreadsheetApp.openById(tablesId);
  const sourceSheetNames = ['Class', 'Race', 'CombatStyles', 'Luck'];
  const destSS = SpreadsheetApp.getActiveSpreadsheet();
  const destSheetName = 'Powers';
  const destSheet = destSS.getSheetByName(destSheetName);

  if (!destSheet) {
    fEndToast();
    fShowMessage('❌ Error', `Could not find the <${destSheetName}> sheet in the current spreadsheet.`);
    return;
  }

  // 3. Prepare for data aggregation and load destination sheet map
  g.DB = {}; // Ensure the namespace for the local DB is fresh
  const { rowTags: destRowTags, colTags: destColTags } = fGetSheetData('DB', destSheetName, destSS);

  // 4. Verify destination column structure
  const columnsToCopy = ['rnd6', 'type', 'subtype', 'tablename', 'source', 'usage', 'action', 'abilityname', 'effect'];
  for (const tag of columnsToCopy) {
    if (destColTags[tag] === undefined) {
      fEndToast();
      fShowMessage('❌ Error', `The <${destSheetName}> sheet must have a column tagged with "${tag}".`);
      return;
    }
  }

  // 5. Clear the destination sheet below the header
  fShowToast('⏳ Clearing old power data...', 'Build Powers');
  const headerRowIndex = destRowTags.header;
  if (headerRowIndex === undefined) {
    fEndToast();
    fShowMessage('❌ Error', `The <${destSheetName}> sheet is missing a "Header" row tag.`);
    return;
  }
  const lastRow = destSheet.getLastRow();
  if (lastRow > headerRowIndex + 1) {
    destSheet.deleteRows(headerRowIndex + 2, lastRow - (headerRowIndex + 1));
  }

  // 6. Process each source sheet and aggregate the data
  const allPowersData = [];
  g.Tbls = {}; // Ensure the namespace exists

  sourceSheetNames.forEach(sourceSheetName => {
    fShowToast(`⏳ Processing <${sourceSheetName}>...`, 'Build Powers');
    const sourceSheet = sourceSS.getSheetByName(sourceSheetName);
    if (!sourceSheet) {
      fShowToast(`⚠️ Could not find sheet: ${sourceSheetName}. Skipping.`, 'Build Powers', 10);
      return; // Continues to the next iteration of forEach
    }

    const { arr: sourceArr, rowTags: sourceRowTags, colTags: sourceColTags } = fGetSheetData('Tbls', sourceSheetName, sourceSS);

    const sourceHeaderIndex = sourceRowTags.header;

    if (sourceHeaderIndex === undefined) {
      fShowToast(`⚠️ No "Header" tag in <${sourceSheetName}>. Skipping.`, 'Build Powers', 10);
      return;
    }

    for (let r = sourceHeaderIndex + 1; r < sourceArr.length; r++) {
      const row = sourceArr[r];
      if (row[sourceColTags.rnd6]) { // Only process rows with an Rnd6 value
        const tableName = row[sourceColTags.tablename];
        const abilityName = row[sourceColTags.abilityname];
        const usage = row[sourceColTags.usage];
        const action = row[sourceColTags.action];
        const effect = row[sourceColTags.effect];

        // This value serves as both the DropDown content and the sort key
        const dropDownValue = `${tableName} - ${abilityName}⚡ (${usage}, ${action}) ➡ ${effect}`;

        const newRow = [
          dropDownValue,
          row[sourceColTags.rnd6],
          row[sourceColTags.type],
          row[sourceColTags.subtype],
          tableName,
          row[sourceColTags.source],
          usage,
          action,
          abilityName,
          effect,
        ];
        allPowersData.push(newRow);
      }
    }
  });


  // 7. Sort the combined array by the first column (the DropDown value)
  fShowToast('⏳ Sorting all powers...', 'Build Powers');
  allPowersData.sort((a, b) => a[0].localeCompare(b[0]));

  // 8. Write the new data to the destination sheet
  if (allPowersData.length > 0) {
    fShowToast(`⏳ Writing ${allPowersData.length} new powers...`, 'Build Powers');
    // Start writing at column 2 (B) to leave column A for row tags
    destSheet.getRange(headerRowIndex + 2, 2, allPowersData.length, allPowersData[0].length).setValues(allPowersData);
  }

  fEndToast();
  fShowMessage('✅ Success', `The <${destSheetName}> sheet has been successfully rebuilt with ${allPowersData.length} powers from all sources.`);
} // End function fBuildPowers