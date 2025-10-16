/* global fShowToast, SpreadsheetApp, fGetSheetData, fShowMessage, fEndToast, fPromptWithInput, g, fGetMasterSheetId, fClearAndWriteData, fActivateSheetByName, fGetVerifiedLocalFile, fGetCodexSpreadsheet, fDeleteTableRow */
/* exported fVerifyIndividualSkills, fVerifySkillSetLists, fBuildSkillSets, fUpdateSkillSetChoices, fFilterSkillSets */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Skill Verification and List Generation
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fApplySkillSetValidations
   Purpose: Reads validation lists from <SkillSetValidationLists> and applies them as dropdowns to the <SkillSets> sheet.
   Assumptions: Running from a Cust sheet.
   Notes: Creates data validation dropdowns to guide user input.
   @returns {void}
*/
function fApplySkillSetValidations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const skillSetsSheet = ss.getSheetByName('SkillSets');
  if (!skillSetsSheet) return;

  const { arr: valArr, rowTags: valRowTags, colTags: valColTags } = fGetSheetData('Cust', 'SkillSetValidationLists', ss);
  const valHeaderRow = valRowTags.header;
  if (valHeaderRow === undefined) return;

  const typeList = valArr.slice(valHeaderRow + 1).map(row => row[valColTags.type]).filter(item => item);
  const subTypeList = valArr.slice(valHeaderRow + 1).map(row => row[valColTags.subtype]).filter(item => item);

  const { rowTags: setsRowTags, colTags: setsColTags } = fGetSheetData('Cust', 'SkillSets', ss);
  const setsHeaderRow = setsRowTags.header;
  const firstDataRow = setsHeaderRow + 2;
  const lastRow = skillSetsSheet.getMaxRows();
  const numRows = lastRow - firstDataRow + 1;

  if (setsHeaderRow === undefined || numRows <= 0) return;

  const typeRule = SpreadsheetApp.newDataValidation().requireValueInList(typeList, true).setAllowInvalid(false).build();
  const subTypeRule = SpreadsheetApp.newDataValidation().requireValueInList(subTypeList, true).setAllowInvalid(false).build();

  skillSetsSheet.getRange(firstDataRow, setsColTags.type + 1, numRows).setDataValidation(typeRule);
  skillSetsSheet.getRange(firstDataRow, setsColTags.subtype + 1, numRows).setDataValidation(subTypeRule);
} // End function fApplySkillSetValidations

/* function fVerifyAndPublishSkillSets
   Purpose: The master workflow for validating and publishing custom skill sets.
   Assumptions: Run from a Cust sheet.
   Notes: This is the gatekeeper for custom skill set data integrity.
   @returns {void}
*/
function fVerifyAndPublishSkillSets() {
  fShowToast('⏳ Verifying skill sets...', '🎓 Verify & Publish');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const validEmojis = ['💪', '🏃', '👁️', '✨'];

  try {
    const { arr: valArr, rowTags: valRowTags, colTags: valColTags } = fGetSheetData('Cust', 'SkillSetValidationLists', ss);
    const valHeaderRow = valRowTags.header;
    if (valHeaderRow === undefined) throw new Error('Could not find <SkillSetValidationLists> or its "Header" tag.');

    const validationLists = {
      typeList: valArr.slice(valHeaderRow + 1).map(row => row[valColTags.type]).filter(Boolean),
      subTypeList: valArr.slice(valHeaderRow + 1).map(row => row[valColTags.subtype]).filter(Boolean),
    };

    const { arr: setsArr, rowTags: setsRowTags, colTags: setsColTags } = fGetSheetData('Cust', 'SkillSets', ss, true);
    const { colTags: destColTags } = fGetSheetData('Cust', 'VerifiedSkillSets', ss);
    const firstDataRowIndex = setsRowTags.header + 1;

    const feedbackData = [];
    const validData = [];
    let passedCount = 0;
    let failedCount = 0;
    const currentUserEmail = Session.getActiveUser().getEmail();

    for (let r = firstDataRowIndex; r < setsArr.length; r++) {
      const row = setsArr[r];
      if (row.every(cell => cell === '')) {
        feedbackData.push(['', '']);
        continue;
      }

      const errors = [];
      const type = row[setsColTags.type];
      if (!type || !validationLists.typeList.includes(type)) errors.push(`Type must be one of: ${validationLists.typeList.join(', ')}.`);

      const subType = row[setsColTags.subtype];
      if (!subType || !validationLists.subTypeList.includes(subType)) errors.push(`SubType must be one of: ${validationLists.subTypeList.join(', ')}.`);

      if (!row[setsColTags.tablename]) errors.push('TableName cannot be empty.');
      if (!row[setsColTags.skillset]) errors.push('SkillSet name cannot be empty.');

      const skillListRaw = row[setsColTags.skilllist] || '';
      const skills = skillListRaw.replace(/,,/g, ',').split(',').map(s => s.trim()).filter(Boolean);

      if (skills.length < 2) errors.push('SkillList must contain at least two comma-separated skills.');
      skills.forEach(skill => {
        if (!validEmojis.some(emoji => skill.endsWith(emoji))) {
          errors.push(`Skill "${skill}" must end with one of: ${validEmojis.join(' ')}.`);
        }
      });

      if (errors.length === 0) {
        passedCount++;
        feedbackData.push(['✅ Passed', '']);
        const skillList = skills.join(', ');
        const dropDownValue = `${row[setsColTags.tablename]} - ${row[setsColTags.skillset]} ➡ ${skillList}`;

        const newValidRow = [];
        newValidRow[destColTags.dropdown] = dropDownValue;
        newValidRow[destColTags.type] = type;
        newValidRow[destColTags.subtype] = subType;
        newValidRow[destColTags.tablename] = row[setsColTags.tablename];
        newValidRow[destColTags.source] = currentUserEmail;
        newValidRow[destColTags.skillset] = row[setsColTags.skillset];
        newValidRow[destColTags.skilllist] = skillList;
        validData.push(newValidRow);
      } else {
        failedCount++;
        feedbackData.push(['❌ Failed', errors.join(' ')]);
      }
    }

    fWriteVerificationResults(ss, 'SkillSets', 'VerifiedSkillSets', feedbackData, validData);

    fEndToast();
    let message = `Verification complete.\n\n✅ ${passedCount} skill sets passed.`;
    if (failedCount > 0) message += `\n❌ ${failedCount} skill sets failed. See 'FailedReason' column for details.`;
    fShowMessage('✅ Verification Complete', message);

  } catch (e) {
    console.error(`❌ CRITICAL ERROR in fVerifyAndPublishSkillSets: ${e.message}\n${e.stack}`);
    fEndToast();
    fShowMessage('❌ Error', `A critical error occurred. Please check the execution logs. Error: ${e.message}`);
  }
} // End function fVerifyAndPublishSkillSets

/* function fDeleteSelectedSkillSets
   Purpose: The master workflow for deleting one or more skill sets from the <SkillSets> sheet.
   Assumptions: Run from a Cust sheet menu.
   Notes: A near-exact copy of fDeleteSelectedPowers.
   @returns {void}
*/
function fDeleteSelectedSkillSets() {
  fShowToast('⏳ Initializing delete...', '🎓 Delete Selected Skill Sets');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'SkillSets';
  const destSheet = ss.getSheetByName(sheetName);

  const { arr, rowTags, colTags } = fGetSheetData('Cust', sheetName, ss, true);
  const headerRow = rowTags.header;

  if (headerRow === undefined) {
    fEndToast();
    fShowMessage('❌ Error', `The <${sheetName}> sheet is missing a "Header" row tag.`);
    return;
  }

  const selectedRows = [];
  for (let r = headerRow + 1; r < arr.length; r++) {
    if (arr[r] && arr[r][colTags.checkbox] === true) {
      selectedRows.push({ row: r + 1, name: arr[r][colTags.skillset] || '' });
    }
  }

  if (selectedRows.length === 0) {
    fEndToast();
    fShowMessage('ℹ️ No Selection', 'Please check the box next to the skill set(s) you wish to delete.');
    return;
  }

  fShowToast('Waiting for your confirmation...', '🎓 Delete Selected Skill Sets');
  const namedItems = selectedRows.filter(p => p.name);
  let promptMessage = '⚠️ Are you sure you wish to permanently DELETE the following?\n';
  let confirmKeyword = 'delete';

  if (namedItems.length > 0) promptMessage += `\n${namedItems.map(p => `- ${p.name}`).join('\n')}\n`;
  promptMessage += '\nThis action cannot be undone.';

  if (selectedRows.length > 1) {
    promptMessage += '\n\nTo confirm, please type DELETE ALL below.';
    confirmKeyword = 'delete all';
  } else {
    promptMessage += '\n\nTo confirm, please type DELETE below.';
  }

  const confirmationText = fPromptWithInput('Confirm Deletion', promptMessage);
  if (confirmationText === null || confirmationText.toLowerCase().trim() !== confirmKeyword) {
    fEndToast();
    fShowMessage('ℹ️ Canceled', 'Deletion has been canceled.');
    return;
  }

  fShowToast('🗑️ Deleting rows...', '🎓 Delete Selected Skill Sets');
  selectedRows.sort((a, b) => b.row - a.row).forEach(item => fDeleteTableRow(destSheet, item.row));

  fEndToast();
  fShowMessage('✅ Deletion Complete', `Successfully deleted ${selectedRows.length} skill set(s).`);
} // End function fDeleteSelectedSkillSets

/* function fGetAllSkillSetTablesList
   Purpose: A helper function to get a definitive, aggregated list of all available skill set tables from the DB.
   Assumptions: None.
   Notes: This is the central source of truth for what skill set tables currently exist.
   @returns {{allSkillSetTables: Array<{tableName: string, source: string}>}} An object containing the aggregated list.
*/
function fGetAllSkillSetTablesList() {
  const dbSkillSetTables = [];
  const customSkillSetTables = [];

  // 1a. Get standard tables from the PLAYER'S LOCAL DB copy.
  const dbFile = fGetVerifiedLocalFile(g.CURRENT_VERSION, 'DB');
  if (dbFile) {
    const sourceSS = SpreadsheetApp.open(dbFile);
    const { arr, rowTags, colTags } = fGetSheetData('DB', 'SkillSets', sourceSS);
    const headerRow = rowTags.header;
    if (headerRow !== undefined) {
      const tableNameCol = colTags.tablename;
      const dbTableNames = [...new Set(arr.slice(headerRow + 1).map(row => row[tableNameCol]).filter(name => name))];
      dbTableNames.forEach(name => dbSkillSetTables.push({ tableName: name, source: 'DB' }));
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
          if (customSS.getSheetByName('VerifiedSkillSets')) {
            const { arr, rowTags, colTags } = fGetSheetData(`Cust_${sourceId}`, 'VerifiedSkillSets', customSS);
            const headerRow = rowTags.header;
            if (headerRow !== undefined) {
              const tableNameCol = colTags.tablename;
              const customTableNames = [...new Set(arr.slice(headerRow + 1).map(row => row[tableNameCol]).filter(name => name))];
              customTableNames.forEach(name => customSkillSetTables.push({ tableName: `Cust - ${name}`, source: sourceName }));
            }
          }
        } catch (e) {
          console.error(`Could not access custom source "${sourceName}" with ID ${sourceId}. Error: ${e}`);
        }
      }
    }
  }


  dbSkillSetTables.sort((a, b) => a.tableName.localeCompare(b.tableName));
  customSkillSetTables.sort((a, b) => a.tableName.localeCompare(b.tableName));
  return { allSkillSetTables: [...dbSkillSetTables, ...customSkillSetTables] };
} // End function fGetAllSkillSetTablesList


/* function fUpdateSkillSetChoices
   Purpose: Updates the <Filter Skill Sets> sheet with a unique list of all TableNames from the PLAYER'S LOCAL DB.
   Assumptions: The user is running this from a Character Sheet.
   Notes: Can be run silently.
   @param {boolean} [isSilent=false] - If true, suppresses the final success message.
   @returns {void}
*/
function fUpdateSkillSetChoices(isSilent = false) {
  fActivateSheetByName('Filter Skill Sets');
  fShowToast('⏳ Syncing skill set tables...', isSilent ? '⚙️ Onboarding' : 'Sync Skill Set Tables');

  const destSS = SpreadsheetApp.getActiveSpreadsheet();
  const destSheet = destSS.getSheetByName('Filter Skill Sets');
  if (!destSheet) {
    if (!isSilent) fEndToast();
    fShowMessage('❌ Error', 'Could not find the <Filter Skill Sets> sheet in this spreadsheet.');
    return;
  }

  const { arr: oldArr, rowTags: oldRowTags, colTags: oldColTags } = fGetSheetData('CS', 'Filter Skill Sets', destSS, true);
  const oldHeaderRow = oldRowTags.header;
  const previouslyChecked = new Set();
  if (oldHeaderRow !== undefined) {
    for (let r = oldHeaderRow + 1; r < oldArr.length; r++) {
      if (oldArr[r][oldColTags.isactive] === true) {
        previouslyChecked.add(oldArr[r][oldColTags.tablename]);
      }
    }
  }

  const { allSkillSetTables } = fGetAllSkillSetTablesList();

  const { rowTags: destRowTags, colTags: destColTags } = fGetSheetData('CS', 'Filter Skill Sets', destSS, true);
  const destHeaderRow = destRowTags.header;
  if (destHeaderRow === undefined) {
    if (!isSilent) fEndToast();
    fShowMessage('❌ Error', 'Could not find a "Header" tag in the <Filter Skill Sets> sheet.');
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

  const newRowCount = allSkillSetTables.length;
  if (newRowCount > 0) {
    if (newRowCount > 1) {
      destSheet.insertRowsAfter(firstDataRow, newRowCount - 1);
    }

    const dataToWrite = allSkillSetTables.map(item => [item.tableName, item.source]);
    destSheet.getRange(firstDataRow, destColTags.tablename + 1, newRowCount, 2).setValues(dataToWrite);

    const newIsActiveCol = destColTags.isactive + 1;
    const newTableNameCol = destColTags.tablename;
    const newData = destSheet.getRange(firstDataRow, newTableNameCol + 1, newRowCount, 1).getValues();

    newData.forEach((row, index) => {
      const tableName = row[0];
      const range = destSheet.getRange(firstDataRow + index, newIsActiveCol);
      if (previouslyChecked.has(tableName)) {
        range.check();
      } else {
        range.insertCheckboxes();
      }
    });
  }

  if (isSilent) {
    fShowToast('✅ Skill set tables synced.', '⚙️ Onboarding');
  } else {
    fEndToast();
    fShowMessage('✅ Success', `The <Filter Skill Sets> sheet has been updated with ${newRowCount} skill set tables.\n\nYour previous selections have been preserved.`);
  }
} // End function fUpdateSkillSetChoices

/* function fPerformSkillSetHealthCheck
   Purpose: A helper to find and remove any stale ("orphaned") skill set tables from the <Filter Skill Sets> sheet.
   Assumptions: None.
   Notes: This is part of the fFilterSkillSets workflow.
   @returns {void}
*/
function fPerformSkillSetHealthCheck() {
  fShowToast('⚕️ Verifying skill set sources...', 'Filter Skill Sets');
  const csSS = SpreadsheetApp.getActiveSpreadsheet();
  const { allSkillSetTables } = fGetAllSkillSetTablesList();
  const validTableNames = new Set(allSkillSetTables.map(t => t.tableName));

  const filterSheet = csSS.getSheetByName('Filter Skill Sets');
  const { arr: choicesArr, rowTags: choicesRowTags, colTags: choicesColTags } = fGetSheetData('CS', 'Filter Skill Sets', csSS, true);
  const choicesHeaderRow = choicesRowTags.header;

  const orphanRows = [];
  for (let r = choicesHeaderRow + 1; r < choicesArr.length; r++) {
    const tableName = choicesArr[r][choicesColTags.tablename];
    if (tableName && !validTableNames.has(tableName)) {
      orphanRows.push({ row: r + 1, name: tableName });
    }
  }

  if (orphanRows.length > 0) {
    fShowToast('🧹 Cleaning up stale entries...', 'Filter Skill Sets');
    orphanRows.sort((a, b) => b.row - a.row).forEach(orphan => {
      fDeleteTableRow(filterSheet, orphan.row);
    });
    const orphanNames = orphanRows.map(o => `- ${o.name}`).join('\n');
    fShowMessage('ℹ️ List Cleaned', `The following skill set tables could no longer be found and have been removed from your list:\n\n${orphanNames}`);
  }
} // End function fPerformSkillSetHealthCheck

/* function fGetSelectedSkillSetTables
   Purpose: A helper to read the <Filter Skill Sets> sheet and return an array of user-selected tables.
   Assumptions: None.
   Notes: This is part of the fFilterSkillSets workflow.
   @returns {Array<object>|null} An array of selected table objects, or null if none are selected or the sheet is invalid.
*/
function fGetSelectedSkillSetTables() {
  const csSS = SpreadsheetApp.getActiveSpreadsheet();
  const { arr, rowTags, colTags } = fGetSheetData('CS', 'Filter Skill Sets', csSS, true);
  const headerRow = rowTags.header;

  const tableNameCol = colTags.tablename;
  const hasContent = arr.slice(headerRow + 1).some(row => row[tableNameCol]);
  if (!hasContent) {
    fEndToast();
    fUpdateSkillSetChoices();
    return null;
  }

  if (headerRow === undefined) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find a "Header" tag in the <Filter Skill Sets> sheet.');
    return null;
  }

  const selectedTables = arr
    .slice(headerRow + 1)
    .filter(row => row[colTags.isactive] === true)
    .map(row => ({ tableName: row[tableNameCol], source: row[colTags.source] }));

  if (selectedTables.length === 0) {
    fEndToast();
    fShowMessage('ℹ️ No Filters Selected', 'Please check one or more boxes on the <Filter Skill Sets> sheet before filtering.');
    return null;
  }

  return selectedTables;
} // End function fGetSelectedSkillSetTables


/* function fFetchAllSkillSetData
   Purpose: A helper to fetch and aggregate all skill set data from the DB.
   Assumptions: None.
   Notes: This is part of the fFilterSkillSets workflow.
   @param {Array<object>} selectedTables - The array of table objects returned by fGetSelectedSkillSetTables.
   @returns {{allSkillSetsData: Array<Array<string>>, dbHeader: Array<string>}|null} An object containing the aggregated data and header, or null on error.
*/
function fFetchAllSkillSetData(selectedTables) {
  fShowToast('Fetching all selected skill sets...', 'Filter Skill Sets');
  let allSkillSetsData = [];
  const codexSS = fGetCodexSpreadsheet();

  const dbFile = fGetVerifiedLocalFile(g.CURRENT_VERSION, 'DB');
  if (!dbFile) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find or restore your local "DB" file to get skill set data from. Please run initial setup.');
    return null;
  }
  const dbSS = SpreadsheetApp.open(dbFile);
  const { arr: allDbSkillSets, rowTags: dbRowTags, colTags: dbColTags } = fGetSheetData('DB', 'SkillSets', dbSS);
  const dbHeader = allDbSkillSets[dbRowTags.header];

  // --- THIS IS THE FIX ---
  // Programmatically correct the header to enforce architectural consistency.
  // This ensures the cache is built correctly even if the DB template is outdated.
  if (dbColTags.effect === undefined) {
    if (dbColTags.skilllist !== undefined) dbHeader[dbColTags.skilllist] = 'Effect';
    else if (dbColTags.skills !== undefined) dbHeader[dbColTags.skills] = 'Effect';
  }
  if (dbColTags.name === undefined && dbColTags.skillset !== undefined) {
    dbHeader[dbColTags.skillset] = 'Name';
  }
  // --- END FIX ---

  const selectedDbTables = selectedTables.filter(t => t.source === 'DB').map(t => t.tableName);
  if (selectedDbTables.length > 0) {
    const dbSkillSets = allDbSkillSets
      .slice(dbRowTags.header + 1)
      .filter(row => selectedDbTables.includes(row[dbColTags.tablename]));
    allSkillSetsData = allSkillSetsData.concat(dbSkillSets);
  }

  // Fetch from Custom Sources
  const selectedCustomTables = selectedTables.filter(t => t.source !== 'DB');
  if (selectedCustomTables.length > 0) {
    const { arr: sourcesArr, colTags: sourcesColTags } = fGetSheetData('Codex', 'Custom Abilities', codexSS, true);
    for (const customTable of selectedCustomTables) {
      const sourceInfo = sourcesArr.find(row => row[sourcesColTags.custabilitiesname] === customTable.source);
      if (sourceInfo) {
        const sourceId = sourceInfo[sourcesColTags.sheetid];
        fShowToast(`Fetching from "${customTable.source}"...`, 'Filter Skill Sets');
        try {
          const customSS = SpreadsheetApp.openById(sourceId);
          const { arr: customSheetSets, rowTags: custRowTags, colTags: custColTags } = fGetSheetData(`Cust_${sourceId}`, 'VerifiedSkillSets', customSS);

          const cleanTableName = customTable.tableName.replace('Cust - ', '');
          const filteredCustomSets = customSheetSets
            .slice(custRowTags.header + 1)
            .filter(row => row[custColTags.tablename] === cleanTableName);

          const mappedCustomSets = filteredCustomSets.map(row => {
            const newRow = [];
            newRow[dbColTags.dropdown] = row[custColTags.dropdown];
            newRow[dbColTags.type] = row[custColTags.type];
            newRow[dbColTags.subtype] = row[custColTags.subtype];
            newRow[dbColTags.tablename] = row[custColTags.tablename];
            newRow[dbColTags.source] = row[custColTags.source];
            newRow[dbColTags.skillset] = row[custColTags.skillset];
            newRow[dbColTags.skilllist] = row[custColTags.skilllist];
            newRow[dbColTags.name] = row[custColTags.skillset];
            newRow[dbColTags.effect] = row[custColTags.skilllist];
            return newRow;
          });
          allSkillSetsData = allSkillSetsData.concat(mappedCustomSets);
        } catch (e) {
          console.error(`Could not access custom source "${customTable.source}". Error: ${e}`);
          fShowMessage('⚠️ Warning', `Could not access the custom source "${customTable.source}". Skipping.`);
        }
      }
    }
  }

  return { allSkillSetsData, dbHeader };
} // End function fFetchAllSkillSetData

/* function fCacheSkillSetData
   Purpose: A helper to write aggregated skill set data to the <SkillSetDataCache> sheet.
   Assumptions: None.
   Notes: This is part of the fFilterSkillSets workflow.
   @param {Array<Array<string>>} allSkillSetsData - The aggregated skill set data.
   @param {Array<string>} dbHeader - The header row for the data.
   @returns {void}
*/
function fCacheSkillSetData(allSkillSetsData, dbHeader) {
  const csSS = SpreadsheetApp.getActiveSpreadsheet();
  const cacheSheet = csSS.getSheetByName('SkillSetDataCache');
  if (!cacheSheet) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find the <SkillSetDataCache> sheet.');
    return;
  }
  cacheSheet.clear();
  if (allSkillSetsData.length > 0) {
    const dataToCache = [dbHeader, ...allSkillSetsData];
    cacheSheet.getRange(1, 1, dataToCache.length, dataToCache[0].length).setValues(dataToCache);
  }
  fShowToast('🎓 Skill set data cached locally.', 'Filter Skill Sets');
} // End function fCacheSkillSetData


/* function fApplySkillSetDropdowns
   Purpose: A helper to build and apply the final data validation dropdowns to the <Game> sheet.
   Assumptions: None.
   Notes: This is part of the fFilterSkillSets workflow.
   @param {Array<Array<string>>} allSkillSetsData - The aggregated skill set data.
   @returns {number} The number of skill sets added to the dropdowns.
*/
function fApplySkillSetDropdowns(allSkillSetsData) {
  const csSS = SpreadsheetApp.getActiveSpreadsheet();
  const { colTags: dbColTags } = fGetSheetData('DB', 'SkillSets', SpreadsheetApp.open(fGetVerifiedLocalFile(g.CURRENT_VERSION, 'DB')));
  const filteredSkillSetList = allSkillSetsData.map(row => row[dbColTags.dropdown]);
  const gameSheet = csSS.getSheetByName('Game');
  if (!gameSheet) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find the <Game> sheet.');
    return 0;
  }

  const { rowTags: gameRowTags, colTags: gameColTags } = fGetSheetData('CS', 'Game', csSS);
  const startRow = gameRowTags.skillsetstart + 1;
  const endRow = gameRowTags.skillsetend + 1;
  const numRows = endRow - startRow + 1;
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(filteredSkillSetList.length > 0 ? filteredSkillSetList : [' '], true).setAllowInvalid(false).build();

  if (gameColTags.skillsetdropdown !== undefined) {
    const colIndex = gameColTags.skillsetdropdown + 1;
    gameSheet.getRange(startRow, colIndex, numRows, 1).setDataValidation(rule);
  }

  return filteredSkillSetList.length;
} // End function fApplySkillSetDropdowns


/* function fFilterSkillSets
   Purpose: Builds custom skill set selection dropdowns on the Character Sheet based on the player's choices in <Filter Skill Sets>.
   Assumptions: The user is running this from a Character Sheet.
   Notes: This is the primary player-facing function for customizing their skill set list.
   @param {boolean} [isSilent=false] - If true, suppresses the final success message.
   @returns {void}
*/
function fFilterSkillSets(isSilent = false) {
  fActivateSheetByName('Filter Skill Sets');

  fPerformSkillSetHealthCheck();

  const selectedTables = fGetSelectedSkillSetTables();
  if (!selectedTables) return;

  const skillSetData = fFetchAllSkillSetData(selectedTables);
  if (!skillSetData) return;

  const { allSkillSetsData, dbHeader } = skillSetData;
  fCacheSkillSetData(allSkillSetsData, dbHeader);

  const finalCount = fApplySkillSetDropdowns(allSkillSetsData);

  if (isSilent) {
    fShowToast('✅ Skill Set dropdowns updated.', '⚙️ Onboarding');
  } else {
    fEndToast();
    fShowMessage('✅ Success!', `Your skill set selection dropdowns have been updated with ${finalCount} skill sets.`);
  }
} // End function fFilterSkillSets

/* function fGetSkillSetSourceData
   Purpose: A helper to fetch, process, and aggregate all skill set data from the master Tables file.
   Assumptions: The 'Tbls' file ID is valid and the <SkillSets> source sheet exists.
   Notes: A helper for the fBuildSkillSets function.
   @param {object} destColTags - The column tag map from the destination <SkillSets> sheet.
   @returns {Array<Array<string>>} A 2D array of the aggregated and processed skill set data.
*/
function fGetSkillSetSourceData(destColTags) {
  const tablesId = fGetMasterSheetId(g.CURRENT_VERSION, 'Tbls');
  if (!tablesId) {
    throw new Error('Could not find the ID for the "Tbls" spreadsheet in the master <Versions> sheet.');
  }

  const sourceSS = SpreadsheetApp.openById(tablesId);
  const sourceSheetName = 'SkillSets';
  fShowToast(`⏳ Processing <${sourceSheetName}>...`, '🎓 Build Skill Sets');
  const sourceSheet = sourceSS.getSheetByName(sourceSheetName);
  if (!sourceSheet) {
    throw new Error(`Could not find the source sheet named "${sourceSheetName}" in the Tables spreadsheet.`);
  }

  g.Tbls = {}; // Ensure a fresh cache namespace
  const { arr: sourceArr, rowTags: sourceRowTags, colTags: sourceColTags } = fGetSheetData('Tbls', sourceSheetName, sourceSS);
  const sourceHeaderIndex = sourceRowTags.header;
  if (sourceHeaderIndex === undefined) {
    throw new Error(`The source <${sourceSheetName}> sheet is missing a "Header" row tag.`);
  }

  const allSkillSetsData = [];
  for (let r = sourceHeaderIndex + 1; r < sourceArr.length; r++) {
    const row = sourceArr[r];
    const type = row[sourceColTags.type];

    // Only process rows that are designated as a "Skill Set"
    if (type === 'Skill Set') {
      const tableName = row[sourceColTags.tablename];
      const skillSet = row[sourceColTags.skillset];
      const skillList = row[sourceColTags.skilllist];
      const dropDownValue = `${tableName} - ${skillSet} ➡ ${skillList}`;

      const newRow = [];
      newRow[destColTags.dropdown] = dropDownValue;
      newRow[destColTags.type] = type;
      newRow[destColTags.subtype] = row[sourceColTags.subtype];
      newRow[destColTags.tablename] = tableName;
      newRow[destColTags.source] = row[sourceColTags.source];
      newRow[destColTags.skillset] = skillSet;
      newRow[destColTags.name] = skillSet;
      newRow[destColTags.effect] = skillList; // --- THIS IS THE FIX ---

      allSkillSetsData.push(newRow);
    }
  }

  // Sort the combined array by the DropDown string
  fShowToast('⏳ Sorting all skill sets...', '🎓 Build Skill Sets');
  allSkillSetsData.sort((a, b) => a[destColTags.dropdown].localeCompare(b[destColTags.dropdown]));

  return allSkillSetsData;
} // End function fGetSkillSetSourceData

/* function fBuildSkillSets
   Purpose: The master function to rebuild the <SkillSets> sheet in the DB file from the master Tables file.
   Assumptions: The user is running this from the DB spreadsheet.
   Notes: This is a destructive and regenerative process.
   @returns {void}
*/
function fBuildSkillSets() {
  fShowToast('⏳ Initializing skill set build...', '🎓 Build Skill Sets');
  const destSheetName = 'SkillSets';
  fActivateSheetByName(destSheetName);

  try {
    const destSS = SpreadsheetApp.getActiveSpreadsheet();
    const destSheet = destSS.getSheetByName(destSheetName);
    if (!destSheet) {
      throw new Error(`Could not find the <${destSheetName}> sheet in the current spreadsheet.`);
    }

    g.DB = {}; // Ensure a fresh cache namespace
    const { colTags: destColTags } = fGetSheetData('DB', destSheetName, destSS, true);

    const allSkillSetData = fGetSkillSetSourceData(destColTags);

    fShowToast(`⏳ Writing ${allSkillSetData.length} new skill sets...`, '🎓 Build Skill Sets');
    fClearAndWriteData(destSheet, allSkillSetData, destColTags);

    fEndToast();
    fShowMessage('✅ Success', `The <${destSheetName}> sheet has been successfully rebuilt with ${allSkillSetData.length} skill sets from the Tables file.`);
  } catch (e) {
    console.error(`❌ CRITICAL ERROR in fBuildSkillSets: ${e.message}\n${e.stack}`);
    fEndToast();
    fShowMessage('❌ Error', `A critical error occurred. Please check the execution logs. Error: ${e.message}`);
  }
} // End function fBuildSkillSets

/* function fVerifySkillSetLists
   Purpose: The master workflow for verifying the skill type emojis within the <SkillSets> sheet.
   Assumptions: Run from a 'Tables' sheet context. The active sheet is <SkillSets>.
   Notes: Iterates through all data rows, splits the comma-separated skill list, and validates each individual skill.
   @returns {void}
*/
function fVerifySkillSetLists() {
  fShowToast('⏳ Verifying all skill sets...', '🎓 Skill Set Verification');
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();

  if (sheetName !== 'SkillSets') {
    fEndToast();
    fShowMessage('⚠️ Warning', 'This function is designed to run only on the <SkillSets> sheet.');
    return;
  }

  try {
    const { arr, rowTags, colTags } = fGetSheetData('Tbls', sheetName, sheet.getParent(), true);
    const headerRow = rowTags.header;
    const skillSetCol = colTags.skillset;
    const skillListCol = colTags.skilllist;

    if (headerRow === undefined || skillSetCol === undefined || skillListCol === undefined) {
      throw new Error(`The <${sheetName}> sheet is missing a required tag (Header, SkillSet, or SkillList).`);
    }

    let correctedCellCount = 0;
    const emojiMap = { '💪': 'Might', '🏃': 'Motion', '👁️': 'Mind', '✨': 'Magic' };
    const validEmojis = Object.keys(emojiMap);

    // Loop through all data rows
    for (let r = headerRow + 1; r < arr.length; r++) {
      const currentRow = r + 1;
      const skillSet = arr[r][skillSetCol];
      const originalSkillList = arr[r][skillListCol];

      // Check the conditions to process a row
      if (skillSet && originalSkillList && originalSkillList.includes(',')) {
        const skills = originalSkillList.split(',').map(s => s.trim());
        const correctedSkills = [];
        let listWasCorrected = false;

        skills.forEach(skill => {
          const correctedSkill = fValidateAndCorrectSkillString(skill, validEmojis, emojiMap);
          if (correctedSkill !== skill) {
            listWasCorrected = true;
          }
          correctedSkills.push(correctedSkill);
        });

        if (listWasCorrected) {
          const newSkillList = correctedSkills.join(', ');
          sheet.getRange(currentRow, skillListCol + 1).setValue(newSkillList);
          correctedCellCount++;
        }
      }
    }

    fEndToast();
    if (correctedCellCount > 0) {
      fShowMessage('✅ Verification Complete', `Found and corrected skills in ${correctedCellCount} skill set(s).`);
    } else {
      fShowMessage('✅ Verification Complete', 'All skill sets are correctly formatted!');
    }
  } catch (e) {
    console.error(`❌ CRITICAL ERROR in fVerifySkillSetLists: ${e.message}\n${e.stack}`);
    fEndToast();
    fShowMessage('❌ Error', `A critical error occurred. Please check the execution logs. Error: ${e.message}`);
  }
} // End function fVerifySkillSetLists

/* function fVerifyIndividualSkills
   Purpose: The master workflow for verifying the skill type emoji in the active sheet.
   Assumptions: Run from a 'Tables' sheet context. The active sheet has a 'Header' row tag and a 'skills' column tag.
   Notes: Iterates through all data rows and uses a helper to validate and correct each skill string.
   @returns {void}
*/
function fVerifyIndividualSkills() {
  fShowToast('⏳ Verifying all skill types...', '🎓 Skill Verification');
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();

  try {
    const { arr, rowTags, colTags } = fGetSheetData('Tbls', sheetName, sheet.getParent(), true);
    const headerRow = rowTags.header;
    const skillsCol = colTags.skills;

    if (headerRow === undefined || skillsCol === undefined) {
      throw new Error(`The <${sheetName}> sheet is missing a "Header" row tag or a "skills" column tag.`);
    }

    let correctedCount = 0;
    const emojiMap = { '💪': 'Might', '🏃': 'Motion', '👁️': 'Mind', '✨': 'Magic' };
    const validEmojis = Object.keys(emojiMap);

    // Loop through all data rows
    for (let r = headerRow + 1; r < arr.length; r++) {
      const currentRow = r + 1;
      const originalString = arr[r][skillsCol];
      if (!originalString) continue; // Skip blank cells

      const correctedString = fValidateAndCorrectSkillString(originalString, validEmojis, emojiMap);

      if (correctedString && correctedString !== originalString) {
        sheet.getRange(currentRow, skillsCol + 1).setValue(correctedString);
        correctedCount++;
      }
    }

    fEndToast();
    if (correctedCount > 0) {
      fShowMessage('✅ Verification Complete', `Found and corrected ${correctedCount} skill type(s).`);
    } else {
      fShowMessage('✅ Verification Complete', 'All skill types are correctly formatted!');
    }
  } catch (e) {
    console.error(`❌ CRITICAL ERROR in fVerifyIndividualSkills: ${e.message}\n${e.stack}`);
    fEndToast();
    fShowMessage('❌ Error', `A critical error occurred. Please check the execution logs. Error: ${e.message}`);
  }
} // End function fVerifyIndividualSkills


/* function fValidateAndCorrectSkillString
   Purpose: Validates a single skill string for the correct emoji and prompts for correction if needed.
   Assumptions: None.
   Notes: A helper for fVerifySkills. Handles auto-correction and re-prompts on invalid input.
   @param {string} skillString - The original string from the 'skills' column.
   @param {Array<string>} validEmojis - An array of the valid emojis.
   @param {object} emojiMap - The map of emojis to their names.
   @returns {string|null} The corrected string, or the original string if no change was needed.
*/
function fValidateAndCorrectSkillString(skillString, validEmojis, emojiMap) {
  // Auto-capitalize every word in the string.
  const capitalizedString = skillString.replace(/\b\w/g, char => char.toUpperCase());
  const foundEmojis = validEmojis.filter(emoji => capitalizedString.includes(emoji));

  // Case 1: Exactly one valid emoji is found.
  if (foundEmojis.length === 1) {
    const emoji = foundEmojis[0];
    const cleanedString = capitalizedString.replace(new RegExp(emoji, 'g'), '').trim();
    const finalString = `${cleanedString}${emoji}`;

    // Auto-correct if the format has changed.
    if (finalString !== skillString) {
      fShowToast(`Fixing format for: "${skillString}"`, '🎓 Skill Verification', 4);
      return finalString;
    }
    // Otherwise, the string is already perfect.
    return skillString;
  }

  // Case 2: Zero or multiple valid emojis are found, requiring user input.
  const choices = validEmojis.map((index, i) => `${i + 1}. ${emojiMap[index]} ${index}`);
  const basePrompt = `The skill has an invalid type:\n\n**${capitalizedString}**\n\nPlease choose the correct type to apply:\n\n${choices.join('\n')}\n\nEnter a number from 1 to ${validEmojis.length}.`;
  let userChoice = null;

  // Loop to re-prompt on invalid input.
  while (true) {
    fShowToast('⚠️ Waiting for your input...', '🎓 Skill Verification');
    const promptMessage = userChoice === null ? basePrompt : `⚠️ Invalid choice. Please try again.\n\n${basePrompt}`;
    userChoice = fPromptWithInput('Correct Skill Type', promptMessage);

    if (userChoice === null) {
      fShowToast('Skipping correction...', '🎓 Skill Verification', 3);
      return skillString; // User canceled.
    }

    const choiceIndex = parseInt(userChoice, 10) - 1;
    if (choiceIndex >= 0 && choiceIndex < validEmojis.length) {
      const correctEmoji = validEmojis[choiceIndex];
      // Remove all old valid emojis before adding the correct one.
      let newString = capitalizedString;
      validEmojis.forEach(emoji => {
        newString = newString.replace(new RegExp(emoji, 'g'), '');
      });
      // Add the correct emoji to the end and trim whitespace.
      return `${newString.trim()}${correctEmoji}`;
    }
    // If input was invalid, the loop will continue and re-prompt.
  }
} // End function fValidateAndCorrectSkillString