/* global g, fGetSheetId, SpreadsheetApp, fBuildTagMaps, fShowMessage, fShowToast */
/* exported fBuildPowers */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Power List Generation
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


/* function fClearAllFilterCheckboxes
   Purpose: Unchecks all 'isactive' checkboxes on a given filter sheet.
   Assumptions: The sheet has a column tagged 'isactive'.
   Notes: A reusable helper for user convenience.
   @param {string} sheetName - The name of the sheet to clear (e.g., 'Filter Powers').
   @returns {void}
*/
function fClearAllFilterCheckboxes(sheetName) {
  fActivateSheetByName(sheetName);
  fShowToast(`⏳ Clearing all selections in <${sheetName}>...`, 'Clear Selections');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    fEndToast();
    fShowMessage('❌ Error', `Could not find the <${sheetName}> sheet.`);
    return;
  }

  const { rowTags, colTags } = fGetSheetData('CS', sheetName, ss, true);
  const headerRow = rowTags.header;
  const isActiveCol = colTags.isactive;

  if (headerRow === undefined || isActiveCol === undefined) {
    fEndToast();
    fShowMessage('❌ Error', `The <${sheetName}> sheet is not tagged correctly.`);
    return;
  }

  const firstDataRow = headerRow + 2;
  const lastRow = sheet.getLastRow();
  const numRows = lastRow - firstDataRow + 1;

  if (numRows > 0) {
    sheet.getRange(firstDataRow, isActiveCol + 1, numRows, 1).uncheck();
  }

  fEndToast();
  fShowMessage('✅ Success', 'All selections have been cleared.\n\n⚠️ You must now select at least one table and run the "Filter..." command or click the green "Refresh" button to update your character\'s dropdowns.');
} // End function fClearAllFilterCheckboxes

/* function fGetAllPowerTablesList
   Purpose: A helper function to get a definitive, aggregated list of all available power tables from DB and Custom sources.
   Assumptions: None.
   Notes: This is the central source of truth for what power tables currently exist.
   @returns {{allPowerTables: Array<{tableName: string, source: string}>}} An object containing the aggregated list.
*/
function fGetAllPowerTablesList() {
  const dbPowerTables = [];
  const customPowerTables = [];

  // 1a. Get standard tables from the PLAYER'S LOCAL DB copy.
  const dbFile = fGetVerifiedLocalFile(g.CURRENT_VERSION, 'DB');
  if (dbFile) {
    const sourceSS = SpreadsheetApp.open(dbFile);
    const { arr, rowTags, colTags } = fGetSheetData('DB', 'Powers', sourceSS);
    const headerRow = rowTags.header;
    if (headerRow !== undefined) {
      const tableNameCol = colTags.tablename;
      const dbTableNames = [...new Set(arr.slice(headerRow + 1).map(row => row[tableNameCol]).filter(name => name))];
      dbTableNames.forEach(name => dbPowerTables.push({ tableName: name, source: 'DB' }));
    }
  }

  // 1b. Get custom tables from all registered sources in the Codex.
  const codexSS = fGetCodexSpreadsheet();
  const { arr: sourcesArr, rowTags: sourcesRowTags, colTags: sourcesColTags } = fGetSheetData('Codex', 'Custom Abilities', codexSS, true);
  const sourcesHeader = sourcesRowTags.header;
  if (sourcesHeader !== undefined) {
    for (let r = sourcesHeader + 1; r < sourcesArr.length; r++) {
      const sourceRow = sourcesArr[r];
      if (sourceRow && sourceRow[sourcesColTags.sheetid]) {
        const sourceId = sourceRow[sourcesColTags.sheetid];
        const sourceName = sourceRow[sourcesColTags.custabilitiesname];
        try {
          const customSS = SpreadsheetApp.openById(sourceId);
          const { arr, rowTags, colTags } = fGetSheetData(`Cust_${sourceId}`, 'VerifiedPowers', customSS);
          const headerRow = rowTags.header;
          if (headerRow !== undefined) {
            const tableNameCol = colTags.tablename;
            const customTableNames = [...new Set(arr.slice(headerRow + 1).map(row => row[tableNameCol]).filter(name => name))];
            customTableNames.forEach(name => customPowerTables.push({ tableName: `Cust - ${name}`, source: sourceName }));
          }
        } catch (e) {
          // Fail silently during the health check
          console.error(`Could not access custom source "${sourceName}" with ID ${sourceId}. Error: ${e}`);
        }
      }
    }
  }

  dbPowerTables.sort((a, b) => a.tableName.localeCompare(b.tableName));
  customPowerTables.sort((a, b) => a.tableName.localeCompare(b.tableName));
  return { allPowerTables: [...dbPowerTables, ...customPowerTables] };
} // End function fGetAllPowerTablesList

/* function fUpdatePowerTablesList
   Purpose: Updates the <Filter Powers> sheet with a unique list of all TableNames from the PLAYER'S LOCAL DB and all registered custom sources.
   Assumptions: The user is running this from a Character Sheet.
   Notes: Aggregates from multiple sources and sorts them into logical groups. Can be run silently.
   @param {boolean} [isSilent=false] - If true, suppresses the final success message.
   @returns {void}
*/
function fUpdatePowerTablesList(isSilent = false) {
  fActivateSheetByName('Filter Powers');
  fShowToast('⏳ Syncing power tables...', isSilent ? '⚙️ Onboarding' : 'Sync Power Tables');

  const destSS = SpreadsheetApp.getActiveSpreadsheet();
  const destSheet = destSS.getSheetByName('Filter Powers');
  if (!destSheet) {
    if (!isSilent) fEndToast();
    fShowMessage('❌ Error', 'Could not find the <Filter Powers> sheet in this spreadsheet.');
    return;
  }

  // --- NEW: Preserve checked state ---
  const { arr: oldArr, rowTags: oldRowTags, colTags: oldColTags } = fGetSheetData('CS', 'Filter Powers', destSS, true);
  const oldHeaderRow = oldRowTags.header;
  const previouslyChecked = new Set();
  if (oldHeaderRow !== undefined) {
    for (let r = oldHeaderRow + 1; r < oldArr.length; r++) {
      if (oldArr[r][oldColTags.isactive] === true) {
        previouslyChecked.add(oldArr[r][oldColTags.tablename]);
      }
    }
  }
  // --- END NEW ---

  const { allPowerTables } = fGetAllPowerTablesList();

  const { rowTags: destRowTags, colTags: destColTags } = fGetSheetData('CS', 'Filter Powers', destSS, true);
  const destHeaderRow = destRowTags.header;
  if (destHeaderRow === undefined) {
    if (!isSilent) fEndToast();
    fShowMessage('❌ Error', 'Could not find a "Header" tag in the <Filter Powers> sheet.');
    return;
  }

  const lastRow = destSheet.getLastRow();
  const firstDataRow = destHeaderRow + 2;
  if (lastRow >= firstDataRow) {
    destSheet.getRange(firstDataRow, 1, lastRow - firstDataRow + 1, destSheet.getMaxColumns()).clearContent();
    if (lastRow > firstDataRow) {
      destSheet.deleteRows(firstDataRow + 1, lastRow - firstDataRow);
    }
  }

  const newRowCount = allPowerTables.length;
  if (newRowCount > 0) {
    if (newRowCount > 1) {
      destSheet.insertRowsAfter(firstDataRow, newRowCount - 1);
      const formatSourceRange = destSheet.getRange(firstDataRow, 1, 1, destSheet.getMaxColumns());
      const formatDestRange = destSheet.getRange(firstDataRow + 1, 1, newRowCount - 1, destSheet.getMaxColumns());
      formatSourceRange.copyTo(formatDestRange, {
        formatOnly: true,
      });
    }

    const dataToWrite = allPowerTables.map(item => {
      const row = [];
      row[destColTags.tablename - 1] = item.tableName;
      row[destColTags.source - 1] = item.source;
      return row;
    });

    destSheet.getRange(firstDataRow, 2, newRowCount, dataToWrite[0].length).setValues(dataToWrite);

    // --- NEW: Re-apply checked state ---
    const newIsActiveCol = destColTags.isactive + 1;
    const newTableNameCol = destColTags.tablename;
    const newData = destSheet.getRange(firstDataRow, newTableNameCol + 1, newRowCount, 1).getValues();

    newData.forEach((row, index) => {
      const tableName = row[0];
      const range = destSheet.getRange(firstDataRow + index, newIsActiveCol);
      if (previouslyChecked.has(tableName)) {
        range.check();
      } else {
        range.insertCheckboxes(); // Ensure even unchecked rows get a box
      }
    });
    // --- END NEW ---
  }

  if (isSilent) {
    fShowToast('✅ Power tables synced.', '⚙️ Onboarding');
  } else {
    fEndToast();
    fShowMessage('✅ Success', `The <Filter Powers> sheet has been updated with ${newRowCount} power tables.\n\nYour previous selections have been preserved.`);
  }
} // End function fUpdatePowerTablesList


/* function fFilterPowers
   Purpose: Builds custom power selection dropdowns on the Character Sheet based on the player's choices in <Filter Powers>, aggregating from DB and Custom sources.
   Assumptions: The user is running this from a Character Sheet.
   Notes: This is the primary player-facing function for customizing their power list. It now also populates a local cache sheet.
   @param {boolean} [isSilent=false] - If true, suppresses the final success message.
   @returns {void}
*/
function fFilterPowers(isSilent = false) {
  fActivateSheetByName('Filter Powers');
  fShowToast('⏳ Filtering power lists...', isSilent ? '⚙️ Onboarding' : 'Filter Powers');

  const csSS = SpreadsheetApp.getActiveSpreadsheet();
  const codexSS = fGetCodexSpreadsheet();

  // --- NEW Health Check Logic ---
  fShowToast('⚕️ Verifying power sources...', isSilent ? '⚙️ Onboarding' : 'Filter Powers');
  const { allPowerTables } = fGetAllPowerTablesList(); // Get a fresh list of ALL valid tables
  const validTableNames = new Set(allPowerTables.map(t => t.tableName));

  const filterSheet = csSS.getSheetByName('Filter Powers');
  const { arr: choicesArr, rowTags: choicesRowTags, colTags: choicesColTags } = fGetSheetData('CS', 'Filter Powers', csSS, true);
  const choicesHeaderRow = choicesRowTags.header;

  const orphanRows = [];
  for (let r = choicesHeaderRow + 1; r < choicesArr.length; r++) {
    const tableName = choicesArr[r][choicesColTags.tablename];
    if (tableName && !validTableNames.has(tableName)) {
      orphanRows.push({ row: r + 1, name: tableName });
    }
  }

  if (orphanRows.length > 0) {
    fShowToast('🧹 Cleaning up stale entries...', isSilent ? '⚙️ Onboarding' : 'Filter Powers');
    orphanRows.sort((a, b) => b.row - a.row).forEach(orphan => {
      fDeleteTableRow(filterSheet, orphan.row);
    });
    const orphanNames = orphanRows.map(o => `- ${o.name}`).join('\n');
    fShowMessage('ℹ️ List Cleaned', `The following power tables could no longer be found and have been removed from your list:\n\n${orphanNames}`);
    // After cleaning, we must re-read the sheet to continue
    fGetSheetData('CS', 'Filter Powers', csSS, true);
  }
  // --- END Health Check ---

  // 1. Read the player's choices from the (potentially cleaned) <Filter Powers> sheet.
  const { arr: finalChoicesArr, rowTags: finalChoicesRowTags, colTags: finalChoicesColTags } = fGetSheetData('CS', 'Filter Powers', csSS);
  const finalChoicesHeaderRow = finalChoicesRowTags.header;

  const tableNameCol = finalChoicesColTags.tablename;
  const hasContent = finalChoicesArr.slice(finalChoicesHeaderRow + 1).some(row => row[tableNameCol]);
  if (!hasContent) {
    if (!isSilent) fEndToast();
    fUpdatePowerTablesList();
    return;
  }

  if (finalChoicesHeaderRow === undefined) {
    if (!isSilent) fEndToast();
    fShowMessage('❌ Error', 'Could not find a "Header" tag in the <Filter Powers> sheet.');
    return;
  }
  const selectedTables = finalChoicesArr
    .slice(finalChoicesHeaderRow + 1)
    .filter(row => row[finalChoicesColTags.isactive] === true)
    .map(row => ({ tableName: row[finalChoicesColTags.tablename], source: row[finalChoicesColTags.source] }));

  if (selectedTables.length === 0) {
    if (!isSilent) fEndToast();
    fShowMessage('ℹ️ No Filters Selected', 'Please check one or more boxes on the <Filter Powers> sheet before filtering.');
    return;
  }

  // 2. Fetch all powers from all selected sources.
  fShowToast('Fetching all selected powers...', isSilent ? '⚙️ Onboarding' : 'Filter Powers');
  let allPowersData = [];
  let dbHeader = [];

  const dbFile = fGetVerifiedLocalFile(g.CURRENT_VERSION, 'DB');
  if (!dbFile) {
    if (!isSilent) fEndToast();
    fShowMessage('❌ Error', 'Could not find or restore your local "DB" file to get power data from. Please run initial setup.');
    return;
  }
  const dbSS = SpreadsheetApp.open(dbFile);
  const { arr: allDbPowers, rowTags: dbRowTags, colTags: dbColTags } = fGetSheetData('DB', 'Powers', dbSS);
  dbHeader = allDbPowers[dbRowTags.header];


  // 2a. Fetch from the local DB if selected
  const selectedDbTables = selectedTables.filter(t => t.source === 'DB').map(t => t.tableName);
  if (selectedDbTables.length > 0) {
    const dbPowers = allDbPowers
      .slice(dbRowTags.header + 1)
      .filter(row => selectedDbTables.includes(row[dbColTags.tablename]));
    allPowersData = allPowersData.concat(dbPowers);
  }

  // 2b. Fetch from Custom Sources
  const selectedCustomTables = selectedTables.filter(t => t.source !== 'DB');
  if (selectedCustomTables.length > 0) {
    const { arr: sourcesArr, colTags: sourcesColTags } = fGetSheetData('Codex', 'Custom Abilities', codexSS, true);
    for (const customTable of selectedCustomTables) {
      const sourceInfo = sourcesArr.find(row => row[sourcesColTags.custabilitiesname] === customTable.source);
      if (sourceInfo) {
        const sourceId = sourceInfo[sourcesColTags.sheetid];
        fShowToast(`Fetching from "${customTable.source}"...`, isSilent ? '⚙️ Onboarding' : 'Filter Powers');
        try {
          const customSS = SpreadsheetApp.openById(sourceId);
          const { arr: customSheetPowers, rowTags: custRowTags, colTags: custColTags } = fGetSheetData(`Cust_${sourceId}`, 'VerifiedPowers', customSS);
          if (dbHeader.length === 0) dbHeader = customSheetPowers[custRowTags.header];

          const cleanTableName = customTable.tableName.replace('Cust - ', '');
          const filteredCustomPowers = customSheetPowers
            .slice(custRowTags.header + 1)
            .filter(row => row[custColTags.tablename] === cleanTableName);

          const mappedCustomPowers = filteredCustomPowers.map(row => {
            const newRow = [];
            newRow[dbColTags.dropdown] = row[custColTags.dropdown];
            newRow[dbColTags.type] = row[custColTags.type];
            newRow[dbColTags.subtype] = row[custColTags.subtype];
            newRow[dbColTags.tablename] = row[custColTags.tablename];
            newRow[dbColTags.source] = row[custColTags.source];
            newRow[dbColTags.usage] = row[custColTags.usage];
            newRow[dbColTags.action] = row[custColTags.action];
            newRow[dbColTags.abilityname] = row[custColTags.abilityname];
            newRow[dbColTags.effect] = row[custColTags.effect];
            return newRow;
          });

          allPowersData = allPowersData.concat(mappedCustomPowers);
        } catch (e) {
          console.error(`Could not access custom source "${customTable.source}". Error: ${e}`);
          fShowMessage('⚠️ Warning', `Could not access the custom source "${customTable.source}". Skipping.`);
        }
      }
    }
  }

  const cacheSheet = csSS.getSheetByName('PowerDataCache');
  if (!cacheSheet) {
    if (!isSilent) fEndToast();
    fShowMessage('❌ Error', 'Could not find the <PowerDataCache> sheet.');
    return;
  }
  cacheSheet.clear();
  if (allPowersData.length > 0) {
    const dataToCache = [dbHeader, ...allPowersData];
    cacheSheet.getRange(1, 1, dataToCache.length, dataToCache[0].length).setValues(dataToCache);
  }
  fShowToast('⚡ Power data cached locally.', isSilent ? '⚙️ Onboarding' : 'Filter Powers');

  const filteredPowerList = allPowersData.map(row => row[dbColTags.dropdown]);
  const gameSheet = csSS.getSheetByName('Game');
  if (!gameSheet) {
    if (!isSilent) fEndToast();
    fShowMessage('❌ Error', 'Could not find the <Game> sheet.');
    return;
  }

  const { rowTags: gameRowTags, colTags: gameColTags } = fGetSheetData('CS', 'Game', csSS);
  const startRow = gameRowTags.powertablestart + 1;
  const endRow = gameRowTags.powertableend + 1;
  const numRows = endRow - startRow + 1;
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(filteredPowerList.length > 0 ? filteredPowerList : [' '], true).setAllowInvalid(false).build();
  
  // --- THIS IS THE FIX ---
  // Use explicit, robust tag checking instead of fragile string matching.
  if (gameColTags.powerdropdown1 !== undefined) {
      const colIndex = gameColTags.powerdropdown1 + 1;
      gameSheet.getRange(startRow, colIndex, numRows, 1).setDataValidation(rule);
  }
  if (gameColTags.powerdropdown2 !== undefined) {
    const colIndex = gameColTags.powerdropdown2 + 1;
    gameSheet.getRange(startRow, colIndex, numRows, 1).setDataValidation(rule);
  }
  // Add more explicit checks here if more dropdowns are added to the <Game> sheet in the future.

  if (isSilent) {
    fShowToast('✅ Power dropdowns updated.', '⚙️ Onboarding');
  } else {
    fEndToast();
    fShowMessage('✅ Success!', `Your power selection dropdowns have been updated with ${filteredPowerList.length} powers.`);
  }
} // End function fFilterPowers

/* function fPerformPowerHealthCheck
   Purpose: A helper to find and remove any stale ("orphaned") power tables from the <Filter Powers> sheet.
   Assumptions: None.
   Notes: This is part of the fFilterPowers workflow.
   @returns {void}
*/
function fPerformPowerHealthCheck() {
  fShowToast('⚕️ Verifying power sources...', 'Filter Powers');
  const csSS = SpreadsheetApp.getActiveSpreadsheet();
  const { allPowerTables } = fGetAllPowerTablesList(); // Get a fresh list of ALL valid tables
  const validTableNames = new Set(allPowerTables.map(t => t.tableName));

  const filterSheet = csSS.getSheetByName('Filter Powers');
  const { arr: choicesArr, rowTags: choicesRowTags, colTags: choicesColTags } = fGetSheetData('CS', 'Filter Powers', csSS, true);
  const choicesHeaderRow = choicesRowTags.header;

  const orphanRows = [];
  for (let r = choicesHeaderRow + 1; r < choicesArr.length; r++) {
    const tableName = choicesArr[r][choicesColTags.tablename];
    if (tableName && !validTableNames.has(tableName)) {
      orphanRows.push({ row: r + 1, name: tableName });
    }
  }

  if (orphanRows.length > 0) {
    fShowToast('🧹 Cleaning up stale entries...', 'Filter Powers');
    orphanRows.sort((a, b) => b.row - a.row).forEach(orphan => {
      fDeleteTableRow(filterSheet, orphan.row);
    });
    const orphanNames = orphanRows.map(o => `- ${o.name}`).join('\n');
    fShowMessage('ℹ️ List Cleaned', `The following power tables could no longer be found and have been removed from your list:\n\n${orphanNames}`);
  }
} // End function fPerformPowerHealthCheck


/* function fGetSelectedPowerTables
   Purpose: A helper to read the <Filter Powers> sheet and return an array of user-selected tables.
   Assumptions: None.
   Notes: This is part of the fFilterPowers workflow.
   @returns {Array<object>|null} An array of selected table objects, or null if none are selected or the sheet is invalid.
*/
function fGetSelectedPowerTables() {
  const csSS = SpreadsheetApp.getActiveSpreadsheet();
  const { arr, rowTags, colTags } = fGetSheetData('CS', 'Filter Powers', csSS, true); // Force refresh after health check
  const headerRow = rowTags.header;

  const tableNameCol = colTags.tablename;
  const hasContent = arr.slice(headerRow + 1).some(row => row[tableNameCol]);
  if (!hasContent) {
    fEndToast();
    fUpdatePowerTablesList(); // No tables listed, so run the sync process for the user.
    return null;
  }

  if (headerRow === undefined) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find a "Header" tag in the <Filter Powers> sheet.');
    return null;
  }

  const selectedTables = arr
    .slice(headerRow + 1)
    .filter(row => row[colTags.isactive] === true)
    .map(row => ({ tableName: row[tableNameCol], source: row[colTags.source] }));

  if (selectedTables.length === 0) {
    fEndToast();
    fShowMessage('ℹ️ No Filters Selected', 'Please check one or more boxes on the <Filter Powers> sheet before filtering.');
    return null;
  }

  return selectedTables;
} // End function fGetSelectedPowerTables


/* function fFetchAllPowerData
   Purpose: A helper to fetch and aggregate all power data from the DB and selected custom sources.
   Assumptions: None.
   Notes: This is part of the fFilterPowers workflow.
   @param {Array<object>} selectedTables - The array of table objects returned by fGetSelectedPowerTables.
   @returns {{allPowersData: Array<Array<string>>, dbHeader: Array<string>}|null} An object containing the aggregated data and header, or null on error.
*/
function fFetchAllPowerData(selectedTables) {
  fShowToast('Fetching all selected powers...', 'Filter Powers');
  let allPowersData = [];
  let dbHeader = [];
  const codexSS = fGetCodexSpreadsheet();

  const dbFile = fGetVerifiedLocalFile(g.CURRENT_VERSION, 'DB');
  if (!dbFile) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find or restore your local "DB" file to get power data from. Please run initial setup.');
    return null;
  }
  const dbSS = SpreadsheetApp.open(dbFile);
  const { arr: allDbPowers, rowTags: dbRowTags, colTags: dbColTags } = fGetSheetData('DB', 'Powers', dbSS);
  dbHeader = allDbPowers[dbRowTags.header];

  // Fetch from the local DB if selected
  const selectedDbTables = selectedTables.filter(t => t.source === 'DB').map(t => t.tableName);
  if (selectedDbTables.length > 0) {
    const dbPowers = allDbPowers
      .slice(dbRowTags.header + 1)
      .filter(row => selectedDbTables.includes(row[dbColTags.tablename]));
    allPowersData = allPowersData.concat(dbPowers);
  }

  // Fetch from Custom Sources
  const selectedCustomTables = selectedTables.filter(t => t.source !== 'DB');
  if (selectedCustomTables.length > 0) {
    const { arr: sourcesArr, colTags: sourcesColTags } = fGetSheetData('Codex', 'Custom Abilities', codexSS, true);
    for (const customTable of selectedCustomTables) {
      const sourceInfo = sourcesArr.find(row => row[sourcesColTags.custabilitiesname] === customTable.source);
      if (sourceInfo) {
        const sourceId = sourceInfo[sourcesColTags.sheetid];
        fShowToast(`Fetching from "${customTable.source}"...`, 'Filter Powers');
        try {
          const customSS = SpreadsheetApp.openById(sourceId);
          const { arr: customSheetPowers, rowTags: custRowTags, colTags: custColTags } = fGetSheetData(`Cust_${sourceId}`, 'VerifiedPowers', customSS);
          if (dbHeader.length === 0) dbHeader = customSheetPowers[custRowTags.header];

          const cleanTableName = customTable.tableName.replace('Cust - ', '');
          const filteredCustomPowers = customSheetPowers
            .slice(custRowTags.header + 1)
            .filter(row => row[custColTags.tablename] === cleanTableName);

          const mappedCustomPowers = filteredCustomPowers.map(row => {
            const newRow = [];
            newRow[dbColTags.dropdown] = row[custColTags.dropdown];
            newRow[dbColTags.type] = row[custColTags.type];
            newRow[dbColTags.subtype] = row[custColTags.subtype];
            newRow[dbColTags.tablename] = row[custColTags.tablename];
            newRow[dbColTags.source] = row[custColTags.source];
            newRow[dbColTags.usage] = row[custColTags.usage];
            newRow[dbColTags.action] = row[custColTags.action];
            newRow[dbColTags.abilityname] = row[custColTags.abilityname];
            newRow[dbColTags.effect] = row[custColTags.effect];
            return newRow;
          });

          allPowersData = allPowersData.concat(mappedCustomPowers);
        } catch (e) {
          console.error(`Could not access custom source "${customTable.source}". Error: ${e}`);
          fShowMessage('⚠️ Warning', `Could not access the custom source "${customTable.source}". Skipping.`);
        }
      }
    }
  }
  return { allPowersData, dbHeader };
} // End function fFetchAllPowerData


/* function fCachePowerData
   Purpose: A helper to write aggregated power data to the <PowerDataCache> sheet.
   Assumptions: None.
   Notes: This is part of the fFilterPowers workflow.
   @param {Array<Array<string>>} allPowersData - The aggregated power data.
   @param {Array<string>} dbHeader - The header row for the data.
   @returns {void}
*/
function fCachePowerData(allPowersData, dbHeader) {
  const csSS = SpreadsheetApp.getActiveSpreadsheet();
  const cacheSheet = csSS.getSheetByName('PowerDataCache');
  if (!cacheSheet) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find the <PowerDataCache> sheet.');
    return;
  }
  cacheSheet.clear();
  if (allPowersData.length > 0) {
    const dataToCache = [dbHeader, ...allPowersData];
    cacheSheet.getRange(1, 1, dataToCache.length, dataToCache[0].length).setValues(dataToCache);
  }
  fShowToast('⚡ Power data cached locally.', 'Filter Powers');
} // End function fCachePowerData


/* function fApplyPowerDropdowns
   Purpose: A helper to build and apply the final data validation dropdowns to the <Game> sheet.
   Assumptions: None.
   Notes: This is part of the fFilterPowers workflow.
   @param {Array<Array<string>>} allPowersData - The aggregated power data.
   @returns {number} The number of powers added to the dropdowns.
*/
function fApplyPowerDropdowns(allPowersData) {
  const csSS = SpreadsheetApp.getActiveSpreadsheet();
  const { colTags: dbColTags } = fGetSheetData('DB', 'Powers', SpreadsheetApp.open(fGetVerifiedLocalFile(g.CURRENT_VERSION, 'DB')));
  const filteredPowerList = allPowersData.map(row => row[dbColTags.dropdown]);
  const gameSheet = csSS.getSheetByName('Game');
  if (!gameSheet) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find the <Game> sheet.');
    return 0;
  }

  const { rowTags: gameRowTags, colTags: gameColTags } = fGetSheetData('CS', 'Game', csSS);
  const startRow = gameRowTags.powertablestart + 1;
  const endRow = gameRowTags.powertableend + 1;
  const numRows = endRow - startRow + 1;
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(filteredPowerList.length > 0 ? filteredPowerList : [' '], true).setAllowInvalid(false).build();

  if (gameColTags.powerdropdown1 !== undefined) {
    const colIndex = gameColTags.powerdropdown1 + 1;
    gameSheet.getRange(startRow, colIndex, numRows, 1).setDataValidation(rule);
  }
  if (gameColTags.powerdropdown2 !== undefined) {
    const colIndex = gameColTags.powerdropdown2 + 1;
    gameSheet.getRange(startRow, colIndex, numRows, 1).setDataValidation(rule);
  }
  
  return filteredPowerList.length;
} // End function fApplyPowerDropdowns











/* function fFilterPowers
   Purpose: Builds custom power selection dropdowns on the Character Sheet based on the player's choices in <Filter Powers>, aggregating from DB and Custom sources.
   Assumptions: The user is running this from a Character Sheet.
   Notes: This is the primary player-facing function for customizing their power list. It now also populates a local cache sheet.
   @param {boolean} [isSilent=false] - If true, suppresses the final success message.
   @returns {void}
*/
function fFilterPowers(isSilent = false) {
  fActivateSheetByName('Filter Powers');
  
  fPerformPowerHealthCheck();

  const selectedTables = fGetSelectedPowerTables();
  if (!selectedTables) return; // Exit if no tables are selected or an error occurred.

  const powerData = fFetchAllPowerData(selectedTables);
  if (!powerData) return; // Exit if there was an error fetching data.

  const { allPowersData, dbHeader } = powerData;
  fCachePowerData(allPowersData, dbHeader);

  const finalCount = fApplyPowerDropdowns(allPowersData);

  if (isSilent) {
    fShowToast('✅ Power dropdowns updated.', '⚙️ Onboarding');
  } else {
    fEndToast();
    fShowMessage('✅ Success!', `Your power selection dropdowns have been updated with ${finalCount} powers.`);
  }
} // End function fFilterPowers