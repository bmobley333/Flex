/* global g, fGetMasterSheetId, SpreadsheetApp, fGetSheetData, fShowToast, fEndToast, fShowMessage */
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
    // 1. Get the ID of the master Tables spreadsheet
    const tablesId = fGetMasterSheetId(g.CURRENT_VERSION, 'Tbls');
    if (!tablesId) {
      fEndToast();
      fShowMessage('❌ Error', 'Could not find the ID for the "Tbls" spreadsheet in the master <Versions> sheet.');
      return;
    }

    // 2. Define source and destination details
    const sourceSS = SpreadsheetApp.openById(tablesId);
    const sourceSheetName = 'Magic Items';
    const sourceSheet = sourceSS.getSheetByName(sourceSheetName);
    const destSS = SpreadsheetApp.getActiveSpreadsheet();
    const destSheet = destSS.getSheetByName(destSheetName);

    if (!sourceSheet) {
      throw new Error(`Could not find the source sheet named "${sourceSheetName}" in the Tables spreadsheet.`);
    }

    // 3. Get fresh data maps for source and destination
    g.DB = {};
    g.Tbls = {};
    const { rowTags: destRowTags, colTags: destColTags } = fGetSheetData('DB', destSheetName, destSS, true);
    const { arr: sourceArr, rowTags: sourceRowTags, colTags: sourceColTags } = fGetSheetData('Tbls', sourceSheetName, sourceSS, true);
    const headerRowIndex = destRowTags.header;
    const sourceHeaderIndex = sourceRowTags.header;

    if (headerRowIndex === undefined || sourceHeaderIndex === undefined) {
      fEndToast();
      fShowMessage('❌ Error', `A "Header" row tag is missing from either the source or destination <Magic Items> sheet.`);
      return;
    }

    // 4. Clear old data from the destination sheet
    fShowToast('⏳ Clearing old magic item data...', '✨ Build Magic Items');
    const lastRow = destSheet.getLastRow();
    const firstDataRow = headerRowIndex + 2;
    if (lastRow >= firstDataRow) {
      destSheet.getRange(firstDataRow, 2, lastRow - firstDataRow + 1, destSheet.getMaxColumns() - 1).clearContent();
      if (lastRow > firstDataRow) {
        destSheet.deleteRows(firstDataRow + 1, lastRow - firstDataRow);
      }
    }

    // 5. Process each source row and aggregate the valid data
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

        allMagicItemsData.push({
          category: category,
          itemName: itemName,
          fullRow: [
            dropDownValue,
            row[sourceColTags.type],
            category,
            row[sourceColTags.tablename],
            row[sourceColTags.source],
            usage,
            action,
            itemName,
            effect,
          ],
        });
      }
    }

    // 6. Sort the aggregated data
    fShowToast('⏳ Sorting all magic items...', '✨ Build Magic Items');
    // --- THIS IS THE FIX ---
    const categoryOrder = ['Minor', 'Lesser', 'Greater', 'Artifact'];
    allMagicItemsData.sort((a, b) => {
      const categoryIndexA = categoryOrder.indexOf(a.category);
      const categoryIndexB = categoryOrder.indexOf(b.category);

      if (categoryIndexA !== categoryIndexB) {
        return categoryIndexA - categoryIndexB;
      }
      return a.itemName.localeCompare(b.itemName);
    });

    const sortedRowData = allMagicItemsData.map(item => item.fullRow);


    // 7. Write the new data to the destination sheet
    const newRowCount = sortedRowData.length;
    if (newRowCount > 0) {
      fShowToast(`⏳ Writing ${newRowCount} new magic items...`, '✨ Build Magic Items');
      if (newRowCount > 1) {
        destSheet.insertRowsAfter(firstDataRow, newRowCount - 1);
        const formatSourceRange = destSheet.getRange(firstDataRow, 1, 1, destSheet.getMaxColumns());
        const formatDestRange = destSheet.getRange(firstDataRow + 1, 1, newRowCount - 1, destSheet.getMaxColumns());
        formatSourceRange.copyTo(formatDestRange, { formatOnly: true });
      }
      destSheet.getRange(firstDataRow, 2, newRowCount, sortedRowData[0].length).setValues(sortedRowData);
    }

    fEndToast();
    fShowMessage('✅ Success', `The <${destSheetName}> sheet has been successfully rebuilt with ${sortedRowData.length} magic items.`);

  } catch (e) {
    console.error(`❌ CRITICAL ERROR in fBuildMagicItems: ${e.message}\n${e.stack}`);
    fEndToast();
    fShowMessage('❌ Error', `A critical error occurred. Please check the execution logs for details. Error: ${e.message}`);
  }
} // End function fBuildMagicItems