/* global g, fGetMasterSheetId, SpreadsheetApp, fGetSheetData, fShowToast, fEndToast, fShowMessage, fActivateSheetByName, fClearAndWriteData */
/* exported fBuildMagicItems */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Magic Item List Generation
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fApplyMagicItemValidations
   Purpose: Reads validation lists from <MagicItemValidationLists> and applies them as dropdowns to the <Magic Items> sheet.
   Assumptions: The script is running from a spreadsheet that contains both a <Magic Items> and a <MagicItemValidationLists> sheet.
   Notes: This function creates data validation dropdowns to guide user input.
   @returns {void}
*/
function fApplyMagicItemValidations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const itemsSheet = ss.getSheetByName('Magic Items');
  if (!itemsSheet) return;

  // 1. Get the validation data
  const { arr: valArr, rowTags: valRowTags, colTags: valColTags } = fGetSheetData('Cust', 'MagicItemValidationLists', ss);
  const valHeaderRow = valRowTags.header;
  if (valHeaderRow === undefined) return;

  // 2. Extract the validation lists into clean arrays
  // --- THIS IS THE FIX ---
  const typeList = valArr.slice(valHeaderRow + 1).map(row => row[valColTags.type]).filter(item => item);
  const subTypeList = valArr.slice(valHeaderRow + 1).map(row => row[valColTags.subtype]).filter(item => item);
  const usageList = valArr.slice(valHeaderRow + 1).map(row => row[valColTags.usage]).filter(item => item);
  const actionList = valArr.slice(valHeaderRow + 1).map(row => row[valColTags.action]).filter(item => item);

  // 3. Get the destination <Magic Items> sheet data to find column locations
  const { rowTags: itemsRowTags, colTags: itemsColTags } = fGetSheetData('Cust', 'Magic Items', ss);
  const itemsHeaderRow = itemsRowTags.header;
  const firstDataRow = itemsHeaderRow + 2;
  const lastRow = itemsSheet.getMaxRows();
  const numRows = lastRow - firstDataRow + 1;

  if (itemsHeaderRow === undefined || numRows <= 0) return;

  // 4. Build and apply the data validation rules
  // --- THIS IS THE FIX ---
  const typeRule = SpreadsheetApp.newDataValidation().requireValueInList(typeList, true).setAllowInvalid(false).build();
  const subTypeRule = SpreadsheetApp.newDataValidation().requireValueInList(subTypeList, true).setAllowInvalid(false).build();
  const usageRule = SpreadsheetApp.newDataValidation().requireValueInList(usageList, true).setAllowInvalid(false).build();
  const actionRule = SpreadsheetApp.newDataValidation().requireValueInList(actionList, true).setAllowInvalid(false).build();

  itemsSheet.getRange(firstDataRow, itemsColTags.type + 1, numRows).setDataValidation(typeRule);
  itemsSheet.getRange(firstDataRow, itemsColTags.subtype + 1, numRows).setDataValidation(subTypeRule);
  itemsSheet.getRange(firstDataRow, itemsColTags.usage + 1, numRows).setDataValidation(usageRule);
  itemsSheet.getRange(firstDataRow, itemsColTags.action + 1, numRows).setDataValidation(actionRule);
} // End function fApplyMagicItemValidations


/* function fValidateMagicItemRow
   Purpose: Validates a single row of data from a <Magic Items> sheet.
   Assumptions: The validation lists have been loaded and passed in.
   Notes: This is the core rules engine for custom magic item validation.
   @param {Array<string>} itemRow - The array of data for a single magic item.
   @param {object} colTags - The column tag map for the sheet.
   @param {object} validationLists - An object containing arrays of valid values (subTypeList, usageList, etc.).
   @returns {{isValid: boolean, errors: Array<string>}} An object indicating if the row is valid and a list of errors.
*/
function fValidateMagicItemRow(itemRow, colTags, validationLists) {
  const errors = [];

  // --- THIS IS THE FIX ---
  // Rule 1: Category must exist and be valid (uses 'subtype' tag)
  const category = itemRow[colTags.subtype];
  if (!category || !validationLists.subTypeList.includes(category)) {
    errors.push(`Category must be one of: ${validationLists.subTypeList.join(', ')}.`);
  }

  // Rule 2: Item Name must exist (uses 'abilityname' tag)
  if (!itemRow[colTags.abilityname]) {
    errors.push('Magic Item\'s Name cannot be empty.');
  }

  // Rule 3: Usage must exist and be valid
  const usage = itemRow[colTags.usage];
  if (!usage || !validationLists.usageList.includes(usage)) {
    errors.push(`Usage must be one of: ${validationLists.usageList.join(', ')}.`);
  }

  // Rule 4: Effect must exist
  if (!itemRow[colTags.effect]) {
    errors.push('Effect cannot be empty.');
  }

  return {
    isValid: errors.length === 0,
    errors: errors,
  };
} // End function fValidateMagicItemRow


/* function fVerifyAndPublishMagicItems
   Purpose: The master workflow for validating and publishing custom magic items.
   Assumptions: Run from a Cust sheet. Reads from <Magic Items>, writes feedback, and copies valid rows to <VerifiedMagicItems>.
   Notes: This is the gatekeeper for ensuring custom magic item data integrity.
   @returns {void}
*/
function fVerifyAndPublishMagicItems() {
  fShowToast('⏳ Verifying magic items...', '✨ Verify & Publish');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('Magic Items');
  const destSheet = ss.getSheetByName('VerifiedMagicItems');
  const currentUserEmail = Session.getActiveUser().getEmail();

  try {
    // 1. Get validation lists
    const { arr: valArr, rowTags: valRowTags, colTags: valColTags } = fGetSheetData('Cust', 'MagicItemValidationLists', ss);
    const valHeaderRow = valRowTags.header;
    if (valHeaderRow === undefined) {
      fEndToast();
      fShowMessage('❌ Error', 'Could not find the <MagicItemValidationLists> sheet or its "Header" tag.');
      return;
    }
    const validationLists = {
      subTypeList: valArr.slice(valHeaderRow + 1).map(row => row[valColTags.subtype]).filter(item => item),
      usageList: valArr.slice(valHeaderRow + 1).map(row => row[valColTags.usage]).filter(item => item),
      actionList: valArr.slice(valHeaderRow + 1).map(row => row[valColTags.action]).filter(item => item),
    };

    // 2. Get data and tags for both sheets
    const { arr: itemsArr, rowTags: itemsRowTags, colTags: itemsColTags } = fGetSheetData('Cust', 'Magic Items', ss, true);
    const { colTags: destColTags } = fGetSheetData('Cust', 'VerifiedMagicItems', ss, true);
    const itemsHeaderRow = itemsRowTags.header;
    const firstDataRowIndex = itemsHeaderRow + 1;

    const feedbackData = [];
    const validItemsData = [];
    let passedCount = 0;
    let failedCount = 0;
    const emojiMap = { Minor: '🍺', Lesser: '🔮', Greater: '🪬', Artifact: '🌀' };

    // 3. Loop, validate, and prepare data
    for (let r = firstDataRowIndex; r < itemsArr.length; r++) {
      const itemRow = itemsArr[r];
      if (itemRow.every(cell => cell === '')) {
        feedbackData.push(['', '']);
        continue;
      }

      const validationResult = fValidateMagicItemRow(itemRow, itemsColTags, validationLists);

      if (validationResult.isValid) {
        passedCount++;
        feedbackData.push(['✅ Passed', '']);

        const category = itemRow[itemsColTags.subtype];
        const itemName = itemRow[itemsColTags.abilityname];
        const usage = itemRow[itemsColTags.usage];
        const action = itemRow[itemsColTags.action];
        const effect = itemRow[itemsColTags.effect];
        const type = itemRow[itemsColTags.type];
        const tableName = itemRow[itemsColTags.tablename];
        const emoji = emojiMap[category] || '✨';
        const dropDownValue = `${category}${emoji} - ${itemName} (${usage}, ${action}) ➡ ${effect}`;

        // --- THIS IS THE FIX ---
        // Build the final row array using the correct DESTINATION tags
        const newValidRow = [];
        newValidRow[destColTags.dropdown] = dropDownValue;
        newValidRow[destColTags.type] = type;
        newValidRow[destColTags.subtype] = category; // Correct destination tag
        newValidRow[destColTags.tablename] = tableName;
        newValidRow[destColTags.source] = currentUserEmail;
        newValidRow[destColTags.usage] = usage;
        newValidRow[destColTags.action] = action;
        newValidRow[destColTags.abilityname] = itemName; // Correct destination tag
        newValidRow[destColTags.effect] = effect;

        validItemsData.push(newValidRow);
      } else {
        failedCount++;
        feedbackData.push(['❌ Failed', validationResult.errors.join(' ')]);
      }
    }

    // 4. Write feedback
    if (feedbackData.length > 0) {
      sourceSheet.getRange(firstDataRowIndex + 1, itemsColTags.verifystatus + 1, feedbackData.length, 2).setValues(feedbackData);
    }

    // 5. Clear and publish valid items
    const destHeaderRow = fGetSheetData('Cust', 'VerifiedMagicItems', ss).rowTags.header;
    const destFirstDataRow = destHeaderRow + 2;
    const lastRow = destSheet.getLastRow();
    if (lastRow >= destFirstDataRow) {
      destSheet.getRange(destFirstDataRow, 1, lastRow - destFirstDataRow + 1, destSheet.getMaxColumns()).clearContent();
    }

    if (validItemsData.length > 0) {
      const outputArr = validItemsData.map(sparseRow => {
        const fullRow = [];
        for (const tag in destColTags) {
          const colIndex = destColTags[tag];
          fullRow[colIndex] = sparseRow[colIndex] || '';
        }
        return fullRow;
      });
      destSheet.getRange(destFirstDataRow, 1, outputArr.length, outputArr[0].length).setValues(outputArr);
    }

    // 6. Display final report
    fEndToast();
    let message = `Verification complete.\n\n✅ ${passedCount} magic items passed and were published.`;
    if (failedCount > 0) {
      message += `\n❌ ${failedCount} magic items failed. Please see the 'FailedReason' column for details.`;
    }
    fShowMessage('✅ Verification Complete', message);

  } catch (e) {
    console.error(`❌ CRITICAL ERROR in fVerifyAndPublishMagicItems: ${e.message}\n${e.stack}`);
    fEndToast();
    fShowMessage('❌ Error', `A critical error occurred. Please check the execution logs. Error: ${e.message}`);
  }
} // End function fVerifyAndPublishMagicItems


/* function fDeleteSelectedMagicItems
   Purpose: The master workflow for deleting one or more items from the active <Magic Items> sheet.
   Assumptions: Run from a Cust sheet menu. The <Magic Items> sheet has a CheckBox column.
   Notes: Includes validation and uses the robust fDeleteTableRow helper.
   @returns {void}
*/
function fDeleteSelectedMagicItems() {
  fShowToast('⏳ Initializing delete...', '✨ Delete Selected Items');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Magic Items';
  const destSheet = ss.getSheetByName(sheetName);

  const { arr, rowTags, colTags } = fGetSheetData('Cust', sheetName, ss, true);
  const headerRow = rowTags.header;

  if (headerRow === undefined) {
    fEndToast();
    fShowMessage('❌ Error', 'The <Magic Items> sheet is missing a "Header" row tag.');
    return;
  }

  // 1. Find all checked rows
  const selectedRows = [];
  for (let r = headerRow + 1; r < arr.length; r++) {
    if (arr[r] && arr[r][colTags.checkbox] === true) {
      selectedRows.push({
        row: r + 1,
        name: arr[r][colTags.abilityname] || '',
      });
    }
  }

  // 2. Validate selection and get confirmation
  if (selectedRows.length === 0) {
    fEndToast();
    fShowMessage('ℹ️ No Selection', 'Please check the box next to the item(s) you wish to delete.');
    return;
  }

  fShowToast('Waiting for your confirmation...', '✨ Delete Selected Items');
  const namedItems = selectedRows.filter(p => p.name);
  const unnamedCount = selectedRows.length - namedItems.length;
  let promptMessage = '⚠️ Are you sure you wish to permanently DELETE the following?\n';
  let confirmKeyword = 'delete';

  if (namedItems.length > 0) {
    promptMessage += `\n${namedItems.map(p => `- ${p.name}`).join('\n')}\n`;
  }
  if (unnamedCount > 0) {
    promptMessage += `\n- ${unnamedCount} unnamed/blank item row${unnamedCount > 1 ? 's' : ''}\n`;
  }
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

  // 3. Delete the spreadsheet rows
  fShowToast('🗑️ Deleting rows...', '✨ Delete Selected Items');
  selectedRows.sort((a, b) => b.row - a.row).forEach(item => {
    fDeleteTableRow(destSheet, item.row);
  });

  fEndToast();
  fShowMessage('✅ Deletion Complete', `Successfully deleted ${selectedRows.length} item(s).`);
} // End function fDeleteSelectedMagicItems

/* function fGetMagicItemSourceData
   Purpose: A helper to fetch, process, and aggregate all magic item data from the master Tables file.
   Assumptions: The 'Tbls' file ID is valid and the 'Magic Items' source sheet exists.
   Notes: This is a helper for the fBuildMagicItems refactor.
   @param {object} destColTags - The column tag map from the destination <Magic Items> sheet.
   @returns {Array<Array<string>>} A 2D array of the aggregated and processed magic item data.
*/
function fGetMagicItemSourceData(destColTags) {
  const tablesId = fGetMasterSheetId(g.CURRENT_VERSION, 'Tbls');
  if (!tablesId) {
    throw new Error('Could not find the ID for the "Tbls" spreadsheet in the master <Versions> sheet.');
  }

  const sourceSS = SpreadsheetApp.openById(tablesId);
  const sourceSheetName = 'Magic Items';
  const sourceSheet = sourceSS.getSheetByName(sourceSheetName);
  if (!sourceSheet) {
    throw new Error(`Could not find the source sheet named "${sourceSheetName}" in the Tables spreadsheet.`);
  }

  g.Tbls = {}; // Ensure a fresh cache namespace
  const { arr: sourceArr, rowTags: sourceRowTags, colTags: sourceColTags } = fGetSheetData('Tbls', sourceSheetName, sourceSS, true);
  const sourceHeaderIndex = sourceRowTags.header;

  if (sourceHeaderIndex === undefined) {
    throw new Error('A "Header" row tag is missing from the source <Magic Items> sheet.');
  }

  fShowToast(`⏳ Processing <${sourceSheetName}>...`, '✨ Build Magic Items');
  const allMagicItemsData = [];
  const emojiMap = { Minor: '🍺', Lesser: '🔮', Greater: '🪬', Artifact: '🌀' };

  for (let r = sourceHeaderIndex + 1; r < sourceArr.length; r++) {
    const row = sourceArr[r];
    const itemName = row[sourceColTags.abilityname];

    if (itemName && itemName.toLowerCase() !== 'item') {
      const category = row[sourceColTags.subtype];
      const usage = row[sourceColTags.usage];
      const action = row[sourceColTags.action];
      const effect = row[sourceColTags.effect];
      const emoji = emojiMap[category] || '✨';
      const dropDownValue = `${category}${emoji} - ${itemName} (${usage}, ${action}) ➡ ${effect}`;

      const newRow = [];
      newRow[destColTags.dropdown] = dropDownValue;
      newRow[destColTags.type] = row[sourceColTags.type];
      newRow[destColTags.subtype] = category;
      newRow[destColTags.tablename] = row[sourceColTags.tablename];
      newRow[destColTags.source] = row[sourceColTags.source];
      newRow[destColTags.usage] = usage;
      newRow[destColTags.action] = action;
      newRow[destColTags.abilityname] = itemName;
      newRow[destColTags.effect] = effect;

      allMagicItemsData.push({
        category: category,
        itemName: itemName,
        fullRow: newRow,
      });
    }
  }

  // Sort the aggregated data
  fShowToast('⏳ Sorting all magic items...', '✨ Build Magic Items');
  const categoryOrder = ['Minor', 'Lesser', 'Greater', 'Artifact'];
  allMagicItemsData.sort((a, b) => {
    const categoryIndexA = categoryOrder.indexOf(a.category);
    const categoryIndexB = categoryOrder.indexOf(b.category);
    if (categoryIndexA !== categoryIndexB) {
      return categoryIndexA - categoryIndexB;
    }
    return a.itemName.localeCompare(b.itemName);
  });

  return allMagicItemsData.map(item => item.fullRow);
} // End function fGetMagicItemSourceData


/* function fBuildMagicItems
   Purpose: The master function to rebuild the <Magic Items> sheet in the DB file from the master Tables file.
   Assumptions: The user is running this from the DB spreadsheet.
   Notes: This is a destructive and regenerative process that reads from the master Tables source sheet.
   @returns {void}
*/
function fBuildMagicItems() {
  fShowToast('⏳ Initializing magic item build...', '✨ Build Magic Items');
  const destSheetName = 'Magic Items';
  fActivateSheetByName(destSheetName);

  try {
    const destSS = SpreadsheetApp.getActiveSpreadsheet();
    const destSheet = destSS.getSheetByName(destSheetName);
    if (!destSheet) {
      throw new Error(`Could not find the <${destSheetName}> sheet in the current spreadsheet.`);
    }

    g.DB = {}; // Ensure a fresh cache namespace
    const { colTags: destColTags } = fGetSheetData('DB', destSheetName, destSS, true);

    const allItemsData = fGetMagicItemSourceData(destColTags);

    fShowToast(`⏳ Writing ${allItemsData.length} new magic items...`, '✨ Build Magic Items');
    fClearAndWriteData(destSheet, allItemsData, destColTags);

    fEndToast();
    fShowMessage('✅ Success', `The <${destSheetName}> sheet has been successfully rebuilt with ${allItemsData.length} magic items.`);
  } catch (e) {
    console.error(`❌ CRITICAL ERROR in fBuildMagicItems: ${e.message}\n${e.stack}`);
    fEndToast();
    fShowMessage('❌ Error', `A critical error occurred. Please check the execution logs for details. Error: ${e.message}`);
  }
} // End function fBuildMagicItems


/* function fGetAllMagicItemsList
   Purpose: A helper to get a definitive, aggregated list of all available magic item TABLES from DB and Custom sources.
   Assumptions: None.
   Notes: This is the central source of truth for what magic item tables currently exist.
   @returns {{allMagicItemTables: Array<{tableName: string, source: string}>}} An object containing the aggregated list.
*/
function fGetAllMagicItemsList() {
  const dbMagicItemTables = [];
  const customMagicItemTables = [];

  // --- THIS IS THE FIX ---
  // Get unique TableNames, not every item.

  // 1a. Get tables from the PLAYER'S LOCAL DB copy.
  const dbFile = fGetVerifiedLocalFile(g.CURRENT_VERSION, 'DB');
  if (dbFile) {
    const sourceSS = SpreadsheetApp.open(dbFile);
    const { arr, rowTags, colTags } = fGetSheetData('DB', 'Magic Items', sourceSS);
    const headerRow = rowTags.header;
    if (headerRow !== undefined) {
      const tableNameCol = colTags.tablename;
      const dbTableNames = [...new Set(arr.slice(headerRow + 1).map(row => row[tableNameCol]).filter(name => name))];
      dbTableNames.forEach(name => dbMagicItemTables.push({ tableName: name, source: 'DB' }));
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
          if (customSS.getSheetByName('VerifiedMagicItems')) {
            const { arr, rowTags, colTags } = fGetSheetData(`Cust_${sourceId}`, 'VerifiedMagicItems', customSS);
            const headerRow = rowTags.header;
            if (headerRow !== undefined) {
              const tableNameCol = colTags.tablename;
              const customTableNames = [...new Set(arr.slice(headerRow + 1).map(row => row[tableNameCol]).filter(name => name))];
              customTableNames.forEach(name => customMagicItemTables.push({ tableName: `Cust - ${name}`, source: sourceName }));
            }
          }
        } catch (e) {
          console.error(`Could not access custom source "${sourceName}" with ID ${sourceId}. Error: ${e}`);
        }
      }
    }
  }

  dbMagicItemTables.sort((a, b) => a.tableName.localeCompare(b.tableName));
  customMagicItemTables.sort((a, b) => a.tableName.localeCompare(b.tableName));
  return { allMagicItemTables: [...dbMagicItemTables, ...customMagicItemTables] };
} // End function fGetAllMagicItemsList

/* function fUpdateMagicItemChoices
   Purpose: Updates the <Filter Magic Items> sheet with a unique list of all TableNames from the DB and all custom sources.
   Assumptions: The user is running this from a Character Sheet.
   Notes: Aggregates from multiple sources and sorts them. Can be run silently.
   @param {boolean} [isSilent=false] - If true, suppresses the final success message.
   @returns {void}
*/
function fUpdateMagicItemChoices(isSilent = false) {
  fActivateSheetByName('Filter Magic Items');
  fShowToast('⏳ Syncing magic item lists...', isSilent ? '⚙️ Onboarding' : '✨ Sync Magic Items');

  const destSS = SpreadsheetApp.getActiveSpreadsheet();
  const destSheet = destSS.getSheetByName('Filter Magic Items');
  if (!destSheet) {
    if (!isSilent) fEndToast();
    fShowMessage('❌ Error', 'Could not find the <Filter Magic Items> sheet in this spreadsheet.');
    return;
  }

  // --- NEW: Preserve checked state ---
  const { arr: oldArr, rowTags: oldRowTags, colTags: oldColTags } = fGetSheetData('CS', 'Filter Magic Items', destSS, true);
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

  const { allMagicItemTables } = fGetAllMagicItemsList();

  const { rowTags: destRowTags, colTags: destColTags } = fGetSheetData('CS', 'Filter Magic Items', destSS, true);
  const destHeaderRow = destRowTags.header;
  if (destHeaderRow === undefined) {
    if (!isSilent) fEndToast();
    fShowMessage('❌ Error', 'Could not find a "Header" tag in the <Filter Magic Items> sheet.');
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

  const newRowCount = allMagicItemTables.length;
  if (newRowCount > 0) {
    if (newRowCount > 1) {
      destSheet.insertRowsAfter(firstDataRow, newRowCount - 1);
    }

    const dataToWrite = allMagicItemTables.map(item => [item.tableName, item.source]);
    destSheet.getRange(firstDataRow, destColTags.tablename + 1, newRowCount, 2).setValues(dataToWrite);

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
        range.insertCheckboxes();
      }
    });
    // --- END NEW ---
  }

  if (isSilent) {
    fShowToast('✅ Magic item tables synced.', '⚙️ Onboarding');
  } else {
    fEndToast();
    fShowMessage('✅ Success', `The <Filter Magic Items> sheet has been updated with ${newRowCount} item tables.\n\nYour previous selections have been preserved.`);
  }
} // End function fUpdateMagicItemChoices


/* function fFilterMagicItems
   Purpose: Builds custom magic item dropdowns on the Character Sheet based on the player's choices in <Filter Magic Items>.
   Assumptions: The user is running this from a Character Sheet.
   Notes: This is the primary player-facing function for customizing their item list. Can be run silently.
   @param {boolean} [isSilent=false] - If true, suppresses the final success message.
   @returns {void}
*/
function fFilterMagicItems(isSilent = false) {
  fActivateSheetByName('Filter Magic Items');
  fShowToast('⏳ Filtering magic items...', isSilent ? '⚙️ Onboarding' : '✨ Filter Magic Items');
  const csSS = SpreadsheetApp.getActiveSpreadsheet();
  const codexSS = fGetCodexSpreadsheet();

  // 1. Read player's choices for which TABLES to include
  const { arr: choicesArr, rowTags: choicesRowTags, colTags: choicesColTags } = fGetSheetData('CS', 'Filter Magic Items', csSS, true);
  const choicesHeaderRow = choicesRowTags.header;

  const tableNameCol = choicesColTags.tablename;
  if (!choicesArr.slice(choicesHeaderRow + 1).some(row => row[tableNameCol])) {
    if (!isSilent) fEndToast();
    fUpdateMagicItemChoices();
    return;
  }

  const selectedTables = choicesArr
    .slice(choicesHeaderRow + 1)
    .filter(row => row[choicesColTags.isactive] === true)
    .map(row => ({ tableName: row[choicesColTags.tablename], source: row[choicesColTags.source] }));

  if (selectedTables.length === 0) {
    if (!isSilent) fEndToast();
    fShowMessage('ℹ️ No Filters Selected', 'Please check one or more boxes on the <Filter Magic Items> sheet before filtering.');
    return;
  }

  // 2. Fetch all item data from all items within the selected TABLES
  fShowToast('Fetching all selected items...', isSilent ? '⚙️ Onboarding' : '✨ Filter Magic Items');
  let allItemsData = [];
  let cacheHeader = [];

  // 2a. Get all data and tags from the DB file one time.
  const dbFile = fGetVerifiedLocalFile(g.CURRENT_VERSION, 'DB');
  const dbSS = SpreadsheetApp.open(dbFile);
  const { arr: allDbItems, rowTags: dbRowTags, colTags: dbColTags } = fGetSheetData('DB', 'Magic Items', dbSS);
  cacheHeader = allDbItems[dbRowTags.header];


  const selectedDbTables = selectedTables.filter(t => t.source === 'DB').map(t => t.tableName);
  if (selectedDbTables.length > 0) {
    const dbItems = allDbItems.filter(row => selectedDbTables.includes(row[dbColTags.tablename]));
    allItemsData = allItemsData.concat(dbItems);
  }

  // 2b. Fetch from Custom Sources
  const selectedCustomTables = selectedTables.filter(t => t.source !== 'DB');
  if (selectedCustomTables.length > 0) {
    const { arr: sourcesArr, colTags: sourcesColTags } = fGetSheetData('Codex', 'Custom Abilities', codexSS, true);
    for (const customTable of selectedCustomTables) {
      const sourceInfo = sourcesArr.find(row => row[sourcesColTags.custabilitiesname] === customTable.source);
      if (sourceInfo) {
        const sourceId = sourceInfo[sourcesColTags.sheetid];
        fShowToast(`Fetching from "${customTable.source}"...`, isSilent ? '⚙️ Onboarding' : 'Filter Magic Items');
        try {
          const customSS = SpreadsheetApp.openById(sourceId);
          const { arr: customSheetItems, rowTags: custRowTags, colTags: custColTags } = fGetSheetData(`Cust_${sourceId}`, 'VerifiedMagicItems', customSS);
          if (cacheHeader.length === 0) cacheHeader = customSheetItems[custRowTags.header];

          const cleanTableName = customTable.tableName.replace('Cust - ', '');
          const filteredCustomItems = customSheetItems
            .slice(custRowTags.header + 1)
            .filter(row => row[custColTags.tablename] === cleanTableName);

          const mappedCustomItems = filteredCustomItems.map(row => {
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

          allItemsData = allItemsData.concat(mappedCustomItems);
        } catch (e) {
          console.error(`Could not access custom source "${customTable.source}". Error: ${e}`);
          fShowMessage('⚠️ Warning', `Could not access the custom source "${customTable.source}". Skipping.`);
        }
      }
    }
  }


  // 3. Populate cache and create dropdowns
  const cacheSheet = csSS.getSheetByName('MagicItemDataCache');
  cacheSheet.clear();

  if (allItemsData.length > 0) {
    cacheSheet.getRange(1, 1, 1, cacheHeader.length).setValues([cacheHeader]);
    cacheSheet.getRange(2, 1, allItemsData.length, allItemsData[0].length).setValues(allItemsData);
  }
  fShowToast('✨ Item data cached locally.', isSilent ? '⚙️ Onboarding' : 'Filter Magic Items');

  const dropDownColIndex = dbColTags.dropdown;
  if (dropDownColIndex === undefined) {
    if (!isSilent) fEndToast();
    fShowMessage('❌ Error', 'Could not find a "dropdown" column tag in the source data.');
    return;
  }

  const filteredItemList = allItemsData.map(row => row[dropDownColIndex]);
  const gameSheet = csSS.getSheetByName('Game');
  const { rowTags: gameRowTags, colTags: gameColTags } = fGetSheetData('CS', 'Game', csSS);
  const startRow = gameRowTags.magicitemtablestart + 1;
  const endRow = gameRowTags.magicitemtableend + 1;
  const numRows = endRow - startRow + 1;
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(filteredItemList.length > 0 ? filteredItemList : [' '], true).setAllowInvalid(false).build();

  if (gameColTags.magicitemdropdown1 !== undefined) {
    const colIndex = gameColTags.magicitemdropdown1 + 1;
    gameSheet.getRange(startRow, colIndex, numRows, 1).setDataValidation(rule);
  }
  if (gameColTags.magicitemdropdown2 !== undefined) {
    const colIndex = gameColTags.magicitemdropdown2 + 1;
    gameSheet.getRange(startRow, colIndex, numRows, 1).setDataValidation(rule);
  }

  if (isSilent) {
    fShowToast('✅ Magic item dropdowns updated.', '⚙️ Onboarding');
  } else {
    fEndToast();
    fShowMessage('✅ Success!', `Your magic item dropdowns have been updated with ${filteredItemList.length} items.`);
  }
} // End function fFilterMagicItems

