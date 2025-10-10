/* global g, fGetSheetId, SpreadsheetApp, fBuildTagMaps, fShowMessage, fShowToast */
/* exported fBuildPowers */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Power List Generation
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fUpdatePowerTablesList
   Purpose: Updates the <Filter Powers> sheet with a unique list of all TableNames from the PLAYER'S LOCAL DB and all registered custom sources.
   Assumptions: The user is running this from a Character Sheet.
   Notes: Aggregates from multiple sources and sorts them into logical groups.
   @returns {void}
*/
function fUpdatePowerTablesList() {
  fActivateSheetByName('Filter Powers');
  fShowToast('⏳ Syncing power tables...', 'Sync Power Tables');

  // --- Part 1: Aggregate All Power Table Information ---
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
      dbTableNames.forEach(name => dbPowerTables.push({
        tableName: name,
        source: 'DB'
      }));
    }
  } else {
    fShowMessage('⚠️ Warning', 'Could not find or restore your local "DB" spreadsheet. Standard powers will be missing.');
  }

  // 1b. Get custom tables from all registered sources in the Codex.
  // (The rest of the function remains the same)
  const codexSS = fGetCodexSpreadsheet();
  const { arr: sourcesArr, rowTags: sourcesRowTags, colTags: sourcesColTags } = fGetSheetData('Codex', 'Custom Abilities', codexSS, true);
  const sourcesHeader = sourcesRowTags.header;

  if (sourcesHeader !== undefined) {
    const sourceIdCol = sourcesColTags.sheetid;
    const sourceNameCol = sourcesColTags.custabilitiesname;
    for (let r = sourcesHeader + 1; r < sourcesArr.length; r++) {
      const sourceRow = sourcesArr[r];
      if (sourceRow && sourceRow[sourceIdCol]) {
        const sourceId = sourceRow[sourceIdCol];
        const sourceName = sourceRow[sourceNameCol];
        fShowToast(`Fetching from "${sourceName}"...`, 'Sync Power Tables');
        try {
          const customSS = SpreadsheetApp.openById(sourceId);
          const { arr, rowTags, colTags } = fGetSheetData(`Cust_${sourceId}`, 'Powers', customSS);
          const headerRow = rowTags.header;
          if (headerRow !== undefined) {
            const tableNameCol = colTags.tablename;
            const customTableNames = [...new Set(arr.slice(headerRow + 1).map(row => row[tableNameCol]).filter(name => name))];
            customTableNames.forEach(name => customPowerTables.push({
              tableName: `Cust - ${name}`,
              source: sourceName
            }));
          }
        } catch (e) {
          console.error(`Could not access custom source "${sourceName}" with ID ${sourceId}. Error: ${e}`);
          fShowMessage('⚠️ Warning', `Could not access the custom source "${sourceName}". It may have been deleted or un-shared. Skipping.`);
        }
      }
    }
  }

  dbPowerTables.sort((a, b) => a.tableName.localeCompare(b.tableName));
  customPowerTables.sort((a, b) => a.tableName.localeCompare(b.tableName));
  const allPowerTables = [...dbPowerTables, ...customPowerTables];

  const destSS = SpreadsheetApp.getActiveSpreadsheet();
  const destSheet = destSS.getSheetByName('Filter Powers');
  if (!destSheet) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find the <Filter Powers> sheet in this spreadsheet.');
    return;
  }

  const { rowTags: destRowTags, colTags: destColTags } = fGetSheetData('CS', 'Filter Powers', destSS, true);
  const destHeaderRow = destRowTags.header;
  if (destHeaderRow === undefined) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find a "Header" tag in the <Filter Powers> sheet.');
    return;
  }

  const lastRow = destSheet.getLastRow();
  const firstDataRow = destHeaderRow + 2;
  if (lastRow >= firstDataRow) {
    destSheet.getRange(firstDataRow, 2, lastRow - firstDataRow + 1, destSheet.getLastColumn() - 1).clearContent();
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
        formatOnly: true
      });
    }

    const dataToWrite = allPowerTables.map(item => {
      const row = [];
      row[destColTags.tablename - 1] = item.tableName;
      row[destColTags.source - 1] = item.source;
      return row;
    });

    destSheet.getRange(firstDataRow, 2, newRowCount, dataToWrite[0].length).setValues(dataToWrite);
    destSheet.getRange(firstDataRow, destColTags.isactive + 1, newRowCount, 1).insertCheckboxes();
  }

  fEndToast();
  fShowMessage('✅ Success', `The <Filter Powers> sheet has been updated with ${newRowCount} power tables.\n\nYou can now check the boxes for the power lists you want to use and then run "Filter Powers" again.`);
} // End function fUpdatePowerTablesList


/* function fFilterPowers
   Purpose: Builds custom power selection dropdowns on the Character Sheet based on the player's choices in <Filter Powers>, aggregating from DB and Custom sources.
   Assumptions: The user is running this from a Character Sheet.
   Notes: This is the primary player-facing function for customizing their power list. It now also populates a local cache sheet.
   @returns {void}
*/
function fFilterPowers() {
  fActivateSheetByName('Filter Powers');
  fShowToast('⏳ Filtering power lists...', 'Filter Powers');

  const csSS = SpreadsheetApp.getActiveSpreadsheet();
  const codexSS = fGetCodexSpreadsheet();

  // 1. Read the player's choices from the <Filter Powers> sheet.
  const { arr: choicesArr, rowTags: choicesRowTags, colTags: choicesColTags } = fGetSheetData('CS', 'Filter Powers', csSS, true);
  const choicesHeaderRow = choicesRowTags.header;

  const tableNameCol = choicesColTags.tablename;
  const hasContent = choicesArr.slice(choicesHeaderRow + 1).some(row => row[tableNameCol]);
  if (!hasContent) {
    fEndToast();
    fUpdatePowerTablesList();
    return;
  }

  if (choicesHeaderRow === undefined) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find a "Header" tag in the <Filter Powers> sheet.');
    return;
  }
  const selectedTables = choicesArr
    .slice(choicesHeaderRow + 1)
    .filter(row => row[choicesColTags.isactive] === true)
    .map(row => ({ tableName: row[choicesColTags.tablename], source: row[choicesColTags.source] }));

  if (selectedTables.length === 0) {
    fEndToast();
    fShowMessage('ℹ️ No Filters Selected', 'Please check one or more boxes on the <Filter Powers> sheet before filtering.');
    return;
  }

  // 2. Fetch all powers from all selected sources.
  fShowToast('Fetching all selected powers...', 'Filter Powers');
  let allPowersData = [];
  let dbHeader = [];

  const dbFile = fGetVerifiedLocalFile(g.CURRENT_VERSION, 'DB');
  if (!dbFile) {
    fEndToast();
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
  // (The rest of the function remains the same)
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
          const { arr: customSheetPowers, rowTags: custRowTags } = fGetSheetData(`Cust_${sourceId}`, 'Powers', customSS);
          if (dbHeader.length === 0) dbHeader = customSheetPowers[custRowTags.header];

          const cleanTableName = customTable.tableName.replace('Cust - ', '');
          const filteredCustomPowers = customSheetPowers
            .slice(custRowTags.header + 1)
            .filter(row => row[dbColTags.tablename] === cleanTableName);

          filteredCustomPowers.forEach(row => {
            const dropDownValue = `Cust - ${row[dbColTags.tablename]} - ${row[dbColTags.abilityname]}⚡ (${row[dbColTags.usage]}, ${row[dbColTags.action]}) ➡ ${row[dbColTags.effect]}`;
            row[dbColTags.dropdown] = dropDownValue;
          });
          allPowersData = allPowersData.concat(filteredCustomPowers);
        } catch (e) {
          console.error(`Could not access custom source "${customTable.source}". Error: ${e}`);
          fShowMessage('⚠️ Warning', `Could not access the custom source "${customTable.source}". Skipping.`);
        }
      }
    }
  }

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

  const filteredPowerList = allPowersData.map(row => row[dbColTags.dropdown]);
  const gameSheet = csSS.getSheetByName('Game');
  if (!gameSheet) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find the <Game> sheet.');
    return;
  }

  const { rowTags: gameRowTags, colTags: gameColTags } = fGetSheetData('CS', 'Game', csSS);
  const startRow = gameRowTags.powertablestart + 1;
  const endRow = gameRowTags.powertableend + 1;
  const numRows = endRow - startRow + 1;
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(filteredPowerList.length > 0 ? filteredPowerList : [' '], true).setAllowInvalid(false).build();
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
  const headerRowIndex = destRowTags.header;

  if (headerRowIndex === undefined) {
    fEndToast();
    fShowMessage('❌ Error', `The <${destSheetName}> sheet is missing a "Header" row tag.`);
    return;
  }

  // 4. Verify destination column structure
  const columnsToCopy = ['type', 'subtype', 'tablename', 'source', 'usage', 'action', 'abilityname', 'effect'];
  for (const tag of columnsToCopy) {
    if (destColTags[tag] === undefined) {
      fEndToast();
      fShowMessage('❌ Error', `The <${destSheetName}> sheet must have a column tagged with "${tag}".`);
      return;
    }
  }

  // 5. Clear the destination sheet using the robust, format-preserving method
  fShowToast('⏳ Clearing old power data...', 'Build Powers');
  const lastRow = destSheet.getLastRow();
  const firstDataRow = headerRowIndex + 2;
  if (lastRow >= firstDataRow) {
    destSheet.getRange(firstDataRow, 2, lastRow - firstDataRow + 1, destSheet.getLastColumn() - 1).clearContent();
    if (lastRow > firstDataRow) {
      destSheet.deleteRows(firstDataRow + 1, lastRow - firstDataRow);
    }
  }

  // 6. Process each source sheet and aggregate the data
  const allPowersData = [];
  g.Tbls = {};

  sourceSheetNames.forEach(sourceSheetName => {
    fShowToast(`⏳ Processing <${sourceSheetName}>...`, 'Build Powers');
    const sourceSheet = sourceSS.getSheetByName(sourceSheetName);
    if (!sourceSheet) {
      fShowToast(`⚠️ Could not find sheet: ${sourceSheetName}. Skipping.`, 'Build Powers', 10);
      return;
    }

    const { arr: sourceArr, rowTags: sourceRowTags, colTags: sourceColTags } = fGetSheetData('Tbls', sourceSheetName, sourceSS);
    const sourceHeaderIndex = sourceRowTags.header;
    if (sourceHeaderIndex === undefined) {
      fShowToast(`⚠️ No "Header" tag in <${sourceSheetName}>. Skipping.`, 'Build Powers', 10);
      return;
    }

    for (let r = sourceHeaderIndex + 1; r < sourceArr.length; r++) {
      const row = sourceArr[r];
      const abilityName = row[sourceColTags.abilityname];

      // --- THIS IS THE FIX ---
      // Only process rows that have a real ability name, not the placeholder "Power" text.
      if (abilityName && abilityName !== 'Power') {
        const tableName = row[sourceColTags.tablename];
        const usage = row[sourceColTags.usage];
        const action = row[sourceColTags.action];
        const effect = row[sourceColTags.effect];
        const dropDownValue = `${tableName} - ${abilityName}⚡ (${usage}, ${action}) ➡ ${effect}`;
        const newRow = [
          dropDownValue,
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

  // 7. Sort the combined array
  fShowToast('⏳ Sorting all powers...', 'Build Powers');
  allPowersData.sort((a, b) => a[0].localeCompare(b[0]));

  // 8. Write the new data
  const newRowCount = allPowersData.length;
  if (newRowCount > 0) {
    fShowToast(`⏳ Writing ${newRowCount} new powers...`, 'Build Powers');
    if (newRowCount > 1) {
      destSheet.insertRowsAfter(firstDataRow, newRowCount - 1);
      const formatSourceRange = destSheet.getRange(firstDataRow, 1, 1, destSheet.getMaxColumns());
      const formatDestRange = destSheet.getRange(firstDataRow + 1, 1, newRowCount - 1, destSheet.getMaxColumns());
      formatSourceRange.copyTo(formatDestRange, { formatOnly: true });
    }
    destSheet.getRange(firstDataRow, 2, newRowCount, allPowersData[0].length).setValues(allPowersData);
  }

  fEndToast();
  fShowMessage('✅ Success', `The <${destSheetName}> sheet has been successfully rebuilt with ${allPowersData.length} powers from all sources.`);
} // End function fBuildPowers