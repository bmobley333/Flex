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
  fShowToast('⏳ Updating power table list...', 'Sync Power Tables');

  // 1. Get the ID for the master DB spreadsheet
  const dbId = fGetMasterSheetId(g.CURRENT_VERSION, 'DB'); // <-- Corrected Line
  if (!dbId) {
    fShowMessage('❌ Error', 'Could not find the master "DB" spreadsheet ID in the master <Versions> sheet.');
    return;
  }

  // 2. Open the source DB <Powers> sheet and get all TableName values
  const sourceSS = SpreadsheetApp.openById(dbId);
  const sourceSheet = sourceSS.getSheetByName('Powers');
  if (!sourceSheet) {
    fShowMessage('❌ Error', 'Could not find the <Powers> sheet in the master DB file.');
    return;
  }

  fLoadSheetToArray('DB', 'Powers', sourceSS);
  fBuildTagMaps('DB', 'Powers');
  const { arr, rowTags, colTags } = g.DB.Powers;
  const startRow = rowTags.header + 1; // Data starts after the header
  const tableNameCol = colTags.tablename;

  const allTableNames = arr.slice(startRow).map(row => row[tableNameCol]);
  const uniqueTableNames = [...new Set(allTableNames)].sort();

  // 3. Get the destination <Choose Powers> sheet and its properties
  const destSS = SpreadsheetApp.getActiveSpreadsheet();
  const destSheet = destSS.getSheetByName('Choose Powers');
  if (!destSheet) {
    fShowMessage('❌ Error', 'Could not find the <Choose Powers> sheet in this spreadsheet.');
    return;
  }

  fLoadSheetToArray('CS', 'Choose Powers', destSS);
  fBuildTagMaps('CS', 'Choose Powers');
  const { rowTags: destRowTags, colTags: destColTags } = g.CS['Choose Powers'];
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

  fShowMessage('✅ Success', `The <Choose Powers> sheet has been updated with ${newRowCount} power tables.`);
} // End function fUpdatePowerTablesList

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
    fShowMessage('❌ Error', 'Could not find the ID for the "Tbls" spreadsheet in the master <Versions> sheet.');
    return;
  }

  // 2. Define source and destination details
  const sourceSS = SpreadsheetApp.openById(tablesId);
  const sourceSheetNames = ['Class', 'Race', 'CombatStyles', 'Luck']; // <-- New array of sources
  const destSS = SpreadsheetApp.getActiveSpreadsheet();
  const destSheetName = 'Powers';
  const destSheet = destSS.getSheetByName(destSheetName);

  if (!destSheet) {
    fShowMessage('❌ Error', `Could not find the <${destSheetName}> sheet in the current spreadsheet.`);
    return;
  }

  // 3. Prepare for data aggregation and load destination sheet map
  g.DB = {}; // Ensure the namespace for the local DB is fresh
  fLoadSheetToArray('DB', destSheetName, destSS);
  fBuildTagMaps('DB', destSheetName);
  const { rowTags: destRowTags, colTags: destColTags } = g.DB[destSheetName];

  // 4. Verify destination column structure
  const columnsToCopy = ['rnd6', 'type', 'subtype', 'tablename', 'source', 'usage', 'action', 'abilityname', 'effect'];
  for (const tag of columnsToCopy) {
    if (destColTags[tag] === undefined) {
      fShowMessage('❌ Error', `The <${destSheetName}> sheet must have a column tagged with "${tag}".`);
      return;
    }
  }

  // 5. Clear the destination sheet below the header
  fShowToast('⏳ Clearing old power data...', 'Build Powers');
  const headerRowIndex = destRowTags.header;
  if (headerRowIndex === undefined) {
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

    fLoadSheetToArray('Tbls', sourceSheetName, sourceSS);
    fBuildTagMaps('Tbls', sourceSheetName);

    const { arr: sourceArr, rowTags: sourceRowTags, colTags: sourceColTags } = g.Tbls[sourceSheetName];
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

  fShowMessage('✅ Success', `The <${destSheetName}> sheet has been successfully rebuilt with ${allPowersData.length} powers from all sources.`);
} // End function fBuildPowers