/* global g, fGetSheetData, SpreadsheetApp, fPromptWithInput, fShowToast, fEndToast, fShowMessage, fGetCodexSpreadsheet, DriveApp, MailApp, Session, Drive, fGetSheetId, fGetOrCreateFolder, fDeleteTableRow */
/* exported fAddOwnCustomAbilitiesSource, fShareMyAbilities, fAddNewCustomSource, fCreateNewCustomList, fRenameCustomList, fDeleteCustomList */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Custom Abilities Management
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* global SpreadsheetApp, fGetSheetData */
/* exported fApplyPowerValidations */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Custom Abilities Management
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fApplyPowerValidations
   Purpose: Reads validation lists from <PowerValidationLists> and applies them as dropdowns to the <Powers> sheet.
   Assumptions: The script is running from a spreadsheet that contains both a <Powers> and a <PowerValidationLists> sheet.
   Notes: This is the definitive function for programmatically creating data validation dropdowns to guide user input.
   @returns {void}
*/
function fApplyPowerValidations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const powersSheet = ss.getSheetByName('Powers');
  if (!powersSheet) return; // Exit silently if the sheet doesn't exist

  // 1. Get the validation data
  const { arr: valArr, rowTags: valRowTags, colTags: valColTags } = fGetSheetData('Cust', 'PowerValidationLists', ss);
  const valHeaderRow = valRowTags.header;
  if (valHeaderRow === undefined) return; // Exit silently if no header

  // 2. Extract the validation lists into clean arrays
  const typeList = valArr.slice(valHeaderRow + 1).map(row => row[valColTags.type]).filter(item => item);
  const subTypeList = valArr.slice(valHeaderRow + 1).map(row => row[valColTags.subtype]).filter(item => item);
  const usageList = valArr.slice(valHeaderRow + 1).map(row => row[valColTags.usage]).filter(item => item);
  const actionList = valArr.slice(valHeaderRow + 1).map(row => row[valColTags.action]).filter(item => item);

  // 3. Get the destination <Powers> sheet data to find column locations
  const { rowTags: powersRowTags, colTags: powersColTags } = fGetSheetData('Cust', 'Powers', ss);
  const powersHeaderRow = powersRowTags.header;
  const firstDataRow = powersHeaderRow + 2; // Start applying validation on the first data row
  const lastRow = powersSheet.getMaxRows();
  const numRows = lastRow - firstDataRow + 1;

  if (powersHeaderRow === undefined || numRows <= 0) return;

  // 4. Build and apply the data validation rules
  const typeRule = SpreadsheetApp.newDataValidation().requireValueInList(typeList, true).setAllowInvalid(false).build();
  const subTypeRule = SpreadsheetApp.newDataValidation().requireValueInList(subTypeList, true).setAllowInvalid(false).build();
  // --- THIS IS THE FIX ---
  const usageRule = SpreadsheetApp.newDataValidation().requireValueInList(usageList, true).setAllowInvalid(false).build();
  const actionRule = SpreadsheetApp.newDataValidation().requireValueInList(actionList, true).setAllowInvalid(false).build();

  powersSheet.getRange(firstDataRow, powersColTags.type + 1, numRows).setDataValidation(typeRule);
  powersSheet.getRange(firstDataRow, powersColTags.subtype + 1, numRows).setDataValidation(subTypeRule);
  powersSheet.getRange(firstDataRow, powersColTags.usage + 1, numRows).setDataValidation(usageRule);
  powersSheet.getRange(firstDataRow, powersColTags.action + 1, numRows).setDataValidation(actionRule);
} // End function fApplyPowerValidations

/* function fValidatePowerRow
   Purpose: Validates a single row of data from a <Powers> sheet.
   Assumptions: The validation lists have been loaded and passed in.
   Notes: This is the core rules engine for custom power validation.
   @param {Array<string>} powerRow - The array of data for a single power.
   @param {object} colTags - The column tag map for the sheet.
   @param {object} validationLists - An object containing arrays of valid values (typeList, subTypeList, actionList).
   @returns {{isValid: boolean, errors: Array<string>}} An object indicating if the row is valid and a list of errors.
*/
function fValidatePowerRow(powerRow, colTags, validationLists) {
  const errors = [];

  // --- NEW, STRICTER RULES ---
  // Rule 1: Type must exist and be valid
  const type = powerRow[colTags.type];
  if (!type || !validationLists.typeList.includes(type)) {
    errors.push(`Type must be one of: ${validationLists.typeList.join(', ')}.`);
  }

  // Rule 2: SubType must exist and be valid
  const subType = powerRow[colTags.subtype];
  if (!subType || !validationLists.subTypeList.includes(subType)) {
    errors.push(`SubType must be one of: ${validationLists.subTypeList.join(', ')}.`);
  }

  // Rule 3: TableName must exist and end with "Powers"
  const tableName = powerRow[colTags.tablename];
  if (!tableName) {
    errors.push('TableName cannot be empty.');
  } else if (!tableName.endsWith('Powers')) {
    errors.push('TableName must end with the word "Powers" (e.g., "My Awesome Powers").');
  }

  // Rule 4: Usage must exist and be valid
  const usage = powerRow[colTags.usage];
  if (!usage || !validationLists.usageList.includes(usage)) {
    errors.push(`Usage must be one of: ${validationLists.usageList.join(', ')}.`);
  }

  // Rule 5: Action must exist and be valid
  const action = powerRow[colTags.action];
  if (!action || !validationLists.actionList.includes(action)) {
    errors.push(`Action must be one of: ${validationLists.actionList.join(', ')}.`);
  }

  // Rule 6: AbilityName must exist
  if (!powerRow[colTags.abilityname]) {
    errors.push('AbilityName cannot be empty.');
  }

  // Rule 7: Effect must exist
  if (!powerRow[colTags.effect]) {
    errors.push('Effect cannot be empty.');
  }

  return {
    isValid: errors.length === 0,
    errors: errors,
  };
} // End function fValidatePowerRow


/* function fVerifyAndPublish
   Purpose: The master workflow for validating and publishing custom powers.
   Assumptions: Run from a Cust sheet. Reads from <Powers>, writes feedback, and copies valid rows to <VerifiedPowers>.
   Notes: This is the definitive gatekeeper for ensuring custom power data integrity.
   @returns {void}
*/
function fVerifyAndPublish() {
  fShowToast('⏳ Verifying abilities...', 'Verify & Publish');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('Powers');
  const destSheet = ss.getSheetByName('VerifiedPowers');
  const currentUserEmail = Session.getActiveUser().getEmail();

  // 1. Get validation lists first
  const { arr: valArr, rowTags: valRowTags, colTags: valColTags } = fGetSheetData('Cust', 'PowerValidationLists', ss);
  const valHeaderRow = valRowTags.header;
  if (valHeaderRow === undefined) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find the <PowerValidationLists> sheet or its "Header" tag.');
    return;
  }
  const validationLists = {
    typeList: valArr.slice(valHeaderRow + 1).map(row => row[valColTags.type]).filter(item => item),
    subTypeList: valArr.slice(valHeaderRow + 1).map(row => row[valColTags.subtype]).filter(item => item),
    usageList: valArr.slice(valHeaderRow + 1).map(row => row[valColTags.usage]).filter(item => item),
    actionList: valArr.slice(valHeaderRow + 1).map(row => row[valColTags.action]).filter(item => item),
  };

  // 2. Get data and tags for both source and destination sheets
  const { arr: powersArr, rowTags: powersRowTags, colTags: powersColTags } = fGetSheetData('Cust', 'Powers', ss, true);
  const { colTags: destColTags } = fGetSheetData('Cust', 'VerifiedPowers', ss);
  const powersHeaderRow = powersRowTags.header;
  const firstDataRowIndex = powersHeaderRow + 1;

  const feedbackData = [];
  const validPowersData = [];
  let passedCount = 0;
  let failedCount = 0;

  // 3. Loop, validate, and prepare feedback/data
  fShowToast('⏳ Validating each power...', 'Verify & Publish');
  for (let r = firstDataRowIndex; r < powersArr.length; r++) {
    const powerRow = powersArr[r];
    const isRowBlank = powerRow.every(cell => cell === '');
    if (isRowBlank) {
      feedbackData.push(['', '']);
      continue;
    }

    const validationResult = fValidatePowerRow(powerRow, powersColTags, validationLists);

    if (validationResult.isValid) {
      passedCount++;
      feedbackData.push(['✅ Passed', '']);

      // --- THIS IS THE NEW LOGIC ---
      // Build the destination row by mapping data using tags
      const newValidRow = [];
      for (const tag in destColTags) {
        const destIndex = destColTags[tag];
        const sourceIndex = powersColTags[tag];

        if (sourceIndex !== undefined) {
          newValidRow[destIndex] = powerRow[sourceIndex];
        }
      }

      // Manually set/override calculated values in the destination row
      newValidRow[destColTags.source] = currentUserEmail;
      newValidRow[destColTags.verifystatus] = '✅ Passed';
      const tableName = newValidRow[destColTags.tablename];
      const abilityName = newValidRow[destColTags.abilityname];
      const usage = newValidRow[destColTags.usage];
      const action = newValidRow[destColTags.action];
      const effect = newValidRow[destColTags.effect];
      newValidRow[destColTags.dropdown] = `${tableName} - ${abilityName}⚡ (${usage}, ${action}) ➡ ${effect}`;

      validPowersData.push(newValidRow);
    } else {
      failedCount++;
      feedbackData.push(['❌ Failed', validationResult.errors.join(' ')]);
    }
  }

  // 4. Write the feedback to the <Powers> sheet
  if (feedbackData.length > 0) {
    sourceSheet.getRange(firstDataRowIndex + 1, powersColTags.verifystatus + 1, feedbackData.length, 2).clearContent();
    sourceSheet.getRange(firstDataRowIndex + 1, powersColTags.verifystatus + 1, feedbackData.length, 2).setValues(feedbackData);
  }

  // 5. Clear and publish valid powers to <VerifiedPowers>
  fShowToast('⏳ Publishing valid powers...', 'Verify & Publish');
  const destHeaderRow = fGetSheetData('Cust', 'VerifiedPowers', ss).rowTags.header;
  const destFirstDataRow = destHeaderRow + 2;
  const lastRow = destSheet.getLastRow();
  if (lastRow >= destFirstDataRow) {
    destSheet.getRange(destFirstDataRow, 1, lastRow - destFirstDataRow + 1, destSheet.getMaxColumns()).clearContent();
  }
  if (validPowersData.length > 0) {
    const destRange = destSheet.getRange(destFirstDataRow, 1, validPowersData.length, destSheet.getMaxColumns());
    const outputArr = destRange.getValues();
    validPowersData.forEach((row, rowIndex) => {
      row.forEach((cell, colIndex) => {
        outputArr[rowIndex][colIndex] = cell;
      });
    });
    destRange.setValues(outputArr);
  }

  // 6. Display the final summary report
  fEndToast();
  const message = `Verification complete.\n\n✅ ${passedCount} powers passed and were published.\n❌ ${failedCount} powers failed. Please see the 'FailedReason' column for details.`;
  fShowMessage('✅ Verification Complete', message);
} // End function fVerifyAndPublish


/* function fDeleteSelectedPowers
   Purpose: The master workflow for deleting one or more powers from the active <Powers> sheet.
   Assumptions: Run from a Cust sheet menu. The <Powers> sheet has a CheckBox column.
   Notes: Includes validation and uses the robust fDeleteTableRow helper to preserve formatting.
   @returns {void}
*/
function fDeleteSelectedPowers() {
  fShowToast('⏳ Initializing delete...', 'Delete Selected Powers');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Powers';
  const destSheet = ss.getSheetByName(sheetName);

  const { arr, rowTags, colTags } = fGetSheetData('Cust', sheetName, ss, true);
  const headerRow = rowTags.header;

  if (headerRow === undefined) {
    fEndToast();
    fShowMessage('❌ Error', 'The <Powers> sheet is missing a "Header" row tag.');
    return;
  }

  // 1. Find all checked rows, regardless of content
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
    fShowMessage('ℹ️ No Selection', 'Please check the box next to the power(s) you wish to delete.');
    return;
  }

  fShowToast('Waiting for your confirmation...', 'Delete Selected Powers');

  const namedPowers = selectedRows.filter(p => p.name);
  const unnamedCount = selectedRows.length - namedPowers.length;
  let promptMessage = '⚠️ Are you sure you wish to permanently DELETE the following?\n';
  let confirmKeyword = 'delete';

  if (namedPowers.length > 0) {
    promptMessage += `\n${namedPowers.map(p => `- ${p.name}`).join('\n')}\n`;
  }
  if (unnamedCount > 0) {
    promptMessage += `\n- ${unnamedCount} unnamed/blank power row${unnamedCount > 1 ? 's' : ''}\n`;
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

  // 3. Delete the spreadsheet rows and track the results
  fShowToast('🗑️ Deleting rows...', 'Delete Selected Powers');
  let clearedCount = 0;
  let deletedCount = 0;
  selectedRows.sort((a, b) => b.row - a.row).forEach(power => {
    const result = fDeleteTableRow(destSheet, power.row);
    if (result === 'cleared') {
      clearedCount++;
    } else {
      deletedCount++;
    }
  });

  // 4. Craft the final, intelligent success message
  fEndToast();
  let successMessage = '';
  if (deletedCount > 0) {
    successMessage += `Successfully deleted ${deletedCount} power(s).`;
  }
  if (clearedCount > 0) {
    if (successMessage) successMessage += '\n\n'; // Add a separator
    successMessage += 'ℹ️ The blank template row was cleared. It remains visible to preserve formatting for when you add new powers.';
  }
  if (!successMessage) {
    successMessage = '✅ Operation complete.'; // Fallback
  }

  fShowMessage('Deletion Complete', successMessage);
} // End function fDeleteSelectedPowers

/* function fDeleteCustomList
   Purpose: The master workflow for deleting one or more player-owned custom ability lists.
   Assumptions: Run from the Codex menu.
   Notes: Includes validation to ensure the user owns all selected lists.
   @returns {void}
*/
function fDeleteCustomList() {
  fShowToast('⏳ Initializing delete...', 'Delete Custom List(s)');
  const ssKey = 'Codex';
  const sheetName = 'Custom Abilities';
  const codexSS = fGetCodexSpreadsheet();
  const currentUser = Session.getActiveUser().getEmail();

  const { arr, rowTags, colTags } = fGetSheetData(ssKey, sheetName, codexSS, true);
  const headerRow = rowTags.header;

  const selectedLists = [];
  for (let r = headerRow + 1; r < arr.length; r++) {
    if (arr[r] && arr[r][colTags.checkbox] === true && arr[r][colTags.custabilitiesname]) { // <-- CHANGE HERE
      selectedLists.push({
        row: r + 1, // 1-based row
        name: arr[r][colTags.custabilitiesname], // <-- CHANGE HERE
        id: arr[r][colTags.sheetid],
        owner: arr[r][colTags.owner],
      });
    }
  }

  // --- Validation ---
  if (selectedLists.length === 0) {
    fEndToast();
    fShowMessage('ℹ️ No Selection', 'Please check the box next to the custom list(s) you wish to delete.');
    return;
  }

  const nonOwnedLists = selectedLists.filter(list => list.owner !== 'Me'); // <-- CHANGE HERE
  if (nonOwnedLists.length > 0) {
    const nonOwnedNames = nonOwnedLists.map(list => list.name).join(', ');
    fEndToast();
    fShowMessage('❌ Permission Denied', `You can only delete custom ability lists that you own. You are not the owner of: ${nonOwnedNames}.`);
    return;
  }
  // --- End Validation ---

  // Confirmation Prompt
  fShowToast('Waiting for your confirmation...', 'Delete Custom List(s)');
  let promptMessage;
  let confirmKeyword;

  if (selectedLists.length === 1) {
    promptMessage = `⚠️ Are you sure you wish to permanently DELETE the custom list "${selectedLists[0].name}"?\n\nThis action cannot be undone.\n\nTo confirm, please type DELETE below.`;
    confirmKeyword = 'delete';
  } else {
    const names = selectedLists.map(c => `- ${c.name}`).join('\n');
    promptMessage = `⚠️ Are you sure you wish to permanently DELETE the following ${selectedLists.length} custom lists?\n\n${names}\n\nThis action cannot be undone.\n\nTo confirm, please type DELETE ALL below.`;
    confirmKeyword = 'delete all';
  }

  const confirmationText = fPromptWithInput('Confirm Deletion', promptMessage);
  if (confirmationText === null || confirmationText.toLowerCase().trim() !== confirmKeyword) {
    fEndToast();
    fShowMessage('ℹ️ Canceled', 'Deletion has been canceled.');
    return;
  }

  // Execute Deletion
  selectedLists.forEach(list => {
    try {
      fShowToast(`🗑️ Trashing file for ${list.name}...`, 'Deleting');
      DriveApp.getFileById(list.id).setTrashed(true);
    } catch (e) {
      console.error(`Could not trash file with ID ${list.id} for list ${list.name}. It may have already been deleted. Error: ${e}`);
    }
  });

  const destSheet = codexSS.getSheetByName(sheetName);
  selectedLists.sort((a, b) => b.row - a.row).forEach(list => {
    fDeleteTableRow(destSheet, list.row);
  });

  fEndToast();
  const deletedNames = selectedLists.map(c => c.name).join(', ');
  fShowMessage('✅ Success', `The following custom list(s) have been deleted:\n\n${deletedNames}`);
} // End function fDeleteCustomList

/* function fRenameCustomList
   Purpose: The master workflow for renaming a player-owned custom ability list.
   Assumptions: Run from the Codex menu.
   Notes: Includes validation to ensure the user is the owner of the list.
   @returns {void}
*/
function fRenameCustomList() {
  fShowToast('⏳ Initializing rename...', 'Rename Custom List');
  const ssKey = 'Codex';
  const sheetName = 'Custom Abilities';
  const codexSS = fGetCodexSpreadsheet();
  const currentUser = Session.getActiveUser().getEmail();

  const { arr, rowTags, colTags } = fGetSheetData(ssKey, sheetName, codexSS, true);
  const headerRow = rowTags.header;

  const selectedLists = [];
  for (let r = headerRow + 1; r < arr.length; r++) {
    if (arr[r] && arr[r][colTags.checkbox] === true) {
      selectedLists.push({
        row: r + 1, // 1-based row
        name: arr[r][colTags.custabilitiesname], // <-- CHANGE HERE
        id: arr[r][colTags.sheetid],
        owner: arr[r][colTags.owner],
        version: g.CURRENT_VERSION, // Assuming current version for simplicity
      });
    }
  }

  // --- Validation ---
  if (selectedLists.length === 0) {
    fEndToast();
    fShowMessage('ℹ️ No Selection', 'Please check the box next to the custom list you wish to rename.');
    return;
  }
  if (selectedLists.length > 1) {
    fEndToast();
    fShowMessage('❌ Error', 'Multiple lists selected. Please select only one list to rename.');
    return;
  }

  const listToRename = selectedLists[0];

  if (listToRename.owner !== 'Me') { // <-- CHANGE HERE
    fEndToast();
    fShowMessage('❌ Permission Denied', 'You can only rename custom ability lists that you own.');
    return;
  }
  // --- End Validation ---

  // Prompt for new name
  const newBaseName = fPromptWithInput('Rename Custom List', `Current Name: ${listToRename.name}\n\nPlease enter the new name for this list:`);
  if (!newBaseName) {
    fEndToast();
    fShowMessage('ℹ️ Canceled', 'Rename operation was canceled.');
    return;
  }

  // Process the new name (strip and re-apply correct version prefix)
  const cleanedName = newBaseName.replace(/^v\d+\s*/, '').trim();
  const finalName = `v${listToRename.version} ${cleanedName}`;

  // Execute the rename
  fShowToast(`Renaming to "${finalName}"...`, 'Rename Custom List');
  try {
    const file = DriveApp.getFileById(listToRename.id);
    file.setName(finalName);

    const nameCell = codexSS.getSheetByName(sheetName).getRange(listToRename.row, colTags.custabilitiesname + 1); // <-- CHANGE HERE
    const url = nameCell.getRichTextValue().getLinkUrl();
    const newLink = SpreadsheetApp.newRichTextValue().setText(finalName).setLinkUrl(url).build();
    nameCell.setRichTextValue(newLink);

    fEndToast();
    fShowMessage('✅ Success', `"${listToRename.name}" has been successfully renamed to "${finalName}".`);
  } catch (e) {
    console.error(`Rename failed. Error: ${e}`);
    fEndToast();
    fShowMessage('❌ Error', 'An error occurred while trying to rename the file. It may have been deleted or you may no longer have permission to edit it.');
  }
} // End function fRenameCustomList

/* function fCreateNewCustomList
   Purpose: Creates a new, named custom ability list from the master template and logs it in the Codex.
   Assumptions: Run from the Codex menu.
   Notes: This is the core workflow for creating a new set of shareable, custom abilities.
   @returns {void}
*/
function fCreateNewCustomList() {
  fShowToast('⏳ Initializing...', 'New Custom List');

  // 1. Get the local Cust template file and destination folder.
  const custTemplateFile = fGetVerifiedLocalFile(g.CURRENT_VERSION, 'Cust');
  if (!custTemplateFile) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find or restore the local master Custom Abilities template.');
    return;
  }

  const customAbilitiesFolder = fGetSubFolder('Custom Abilities');
  if (!customAbilitiesFolder) {
    fEndToast(); // fGetSubFolder shows its own error message.
    return;
  }

  // 2. Copy the template.
  fShowToast('Copying template...', 'New Custom List');
  const newCustFile = custTemplateFile.makeCopy(customAbilitiesFolder);
  const newCustSS = SpreadsheetApp.openById(newCustFile.getId());
  fEmbedCodexId(newCustSS);

  // 3. Prompt for a name.
  const listName = fPromptWithInput('Name Your List', 'Please enter a name for your new custom ability list (e.g., "My Homebrew Powers"):');
  if (!listName) {
    newCustFile.setTrashed(true); // Clean up if the user cancels
    fEndToast();
    fShowMessage('ℹ️ Canceled', 'Creation of new custom ability list was canceled.');
    return;
  }

  // Apply versioning to the name.
  const versionedListName = `v${g.CURRENT_VERSION} ${listName.replace(/^v\d+\s*/, '').trim()}`;
  newCustFile.setName(versionedListName);

  // 4. Log the new list in the Codex's <Custom Abilities> sheet.
  fShowToast('Logging new list in your Codex...', 'New Custom List');
  const ssKey = 'Codex';
  const sheetName = 'Custom Abilities';
  const codexSS = fGetCodexSpreadsheet();
  const destSheet = codexSS.getSheetByName(sheetName);
  const { arr, rowTags, colTags } = fGetSheetData(ssKey, sheetName, codexSS, true);
  const headerRow = rowTags.header;
  const lastRow = destSheet.getLastRow();
  let targetRow;

  const dataToWrite = [];
  dataToWrite[colTags.sheetid - 1] = newCustFile.getId();
  dataToWrite[colTags.custabilitiesname - 1] = versionedListName;
  dataToWrite[colTags.owner - 1] = 'Me';

  const firstDataRowIndex = headerRow + 1;
  const templateRow = firstDataRowIndex + 1;
  const nameCol = colTags.custabilitiesname;

  if (arr.length <= firstDataRowIndex || !arr[firstDataRowIndex][nameCol]) {
    targetRow = templateRow;
  } else {
    targetRow = lastRow + 1;
    destSheet.insertRowsAfter(lastRow, 1);
    const formatSourceRange = destSheet.getRange(templateRow, 1, 1, destSheet.getMaxColumns());
    const formatDestRange = destSheet.getRange(targetRow, 1, 1, destSheet.getMaxColumns());
    formatSourceRange.copyTo(formatDestRange, { formatOnly: true });
  }

  const targetRange = destSheet.getRange(targetRow, 2, 1, dataToWrite.length);
  targetRange.setValues([dataToWrite]);

  const link = SpreadsheetApp.newRichTextValue().setText(versionedListName).setLinkUrl(newCustFile.getUrl()).build();
  destSheet.getRange(targetRow, colTags.custabilitiesname + 1).setRichTextValue(link);

  fEndToast();
  fShowMessage('✅ Success', `Your new custom ability list "${listName}" has been created and added to your Codex.`);
} // End function fCreateNewCustomList

/* function fShareCustomLists
   Purpose: Orchestrates the workflow for sharing one or more player-owned custom ability lists.
   Assumptions: Run from the Codex menu. The advanced Drive API service must be enabled.
   Notes: Grants viewer permission and sends a custom notification email.
   @returns {void}
*/
function fShareCustomLists() {
  fShowToast('⏳ Initializing share...', 'Share Custom Lists');
  const ssKey = 'Codex';
  const sheetName = 'Custom Abilities';
  const codexSS = fGetCodexSpreadsheet();
  const currentUser = Session.getActiveUser().getEmail();

  // 1. Find all checked lists
  const { arr, rowTags, colTags } = fGetSheetData(ssKey, sheetName, codexSS, true);
  const headerRow = rowTags.header;

  const selectedLists = [];
  for (let r = headerRow + 1; r < arr.length; r++) {
    if (arr[r] && arr[r][colTags.checkbox] === true && arr[r][colTags.custabilitiesname]) { // <-- CHANGE HERE
      selectedLists.push({
        name: arr[r][colTags.custabilitiesname], // <-- CHANGE HERE
        id: arr[r][colTags.sheetid],
        owner: arr[r][colTags.owner],
      });
    }
  }

  // 2. --- Validation ---
  if (selectedLists.length === 0) {
    fEndToast();
    fShowMessage('ℹ️ No Selection', 'Please check the box next to the custom list(s) you wish to share.');
    return;
  }

  const nonOwnedLists = selectedLists.filter(list => list.owner !== 'Me'); // <-- CHANGE HERE
  if (nonOwnedLists.length > 0) {
    const nonOwnedNames = nonOwnedLists.map(list => list.name).join(', ');
    fEndToast();
    fShowMessage('❌ Permission Denied', `You can only share custom ability lists that you own. You are not the owner of: ${nonOwnedNames}.`);
    return;
  }
  // --- End Validation ---

  // 3. Prompt for the recipient's email address
  const listNamesForPrompt = selectedLists.map(c => `- ${c.name}`).join('\n');
  const promptMessage = `You are about to share the following ${selectedLists.length} list(s):\n\n${listNamesForPrompt}\n\nEnter the email address of the player you want to share these files with:`;
  const email = fPromptWithInput('Share Custom Lists', promptMessage);

  if (!email) {
    fEndToast();
    fShowMessage('ℹ️ Canceled', 'Sharing was canceled.');
    return;
  }

  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailRegex.test(email)) {
    fEndToast();
    fShowMessage('❌ Invalid Email', 'The email address you entered does not appear to be valid. Please try again.');
    return;
  }

  // 4. Grant permissions and send a consolidated email
  try {
    fShowToast(`Sharing ${selectedLists.length} file(s) with ${email}...`, 'Share Custom Lists');
    selectedLists.forEach(list => {
      const permissionResource = {
        role: 'reader',
        type: 'user',
        emailAddress: email,
      };
      Drive.Permissions.create(permissionResource, list.id, {
        sendNotificationEmail: false,
      });
    });

    fShowToast('Sending notification email...', 'Share Custom Lists');
    const subject = `Flex TTRPG: ${selectedLists.length} custom list(s) have been shared with you!`;
    const listDetailsForEmail = selectedLists.map(list => `Name: ${list.name}.    ID below:\n${list.id}`).join('\n\n');
    const body = `The player ${currentUser} has shared the following Flex Custom Abilities sheet(s) with you.\n\n` +
      `To add them, open your Player's Codex, go to "*** Flex ***" > "Custom Abilities" > "Add Sheet From ID", and paste the ID for each sheet when prompted (For multiple sheets, repeat this for each ID below).\n\n` +
      `----------------------------------------\n\n` +
      `${listDetailsForEmail}\n\n` +
      `----------------------------------------`;
    MailApp.sendEmail(email, subject, body);

    fEndToast();
    fShowMessage('✅ Success!', `Your custom list(s) have been successfully shared with ${email}.`);
  } catch (e) {
    console.error(`Sharing failed. Error: ${e}`);
    fEndToast();
    fShowMessage('❌ Error', 'An error occurred while trying to share the file(s). Please ensure the advanced Drive API is enabled for the Codex project.');
  }
} // End function fShareCustomLists


/* function fAddNewCustomSource
   Purpose: The master workflow for adding a new, external custom ability source to the Codex.
   Assumptions: Run from the Codex menu.
   Notes: Includes validation, permission checks, and user prompts.
   @returns {void}
*/
function fAddNewCustomSource() {
  fShowToast('⏳ Initializing...', 'Add New Source');

  // 1. Prompt for the Sheet ID
  const sourceId = fPromptWithInput('Add Custom Source', 'Please enter the Google Sheet ID of the custom abilities file you want to add:');
  if (!sourceId) {
    fEndToast();
    fShowMessage('ℹ️ Canceled', 'Operation was canceled.');
    return;
  }

  // 2. Verify the ID and permissions
  let sourceSS;
  let ownerEmail;
  const currentUser = Session.getActiveUser().getEmail();
  try {
    fShowToast('Verifying ID and permissions...', 'Add New Source');
    sourceSS = SpreadsheetApp.openById(sourceId);
    ownerEmail = sourceSS.getOwner() ? sourceSS.getOwner().getEmail() : 'Unknown';
  } catch (e) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not access the spreadsheet. Please check that the ID is correct and that the owner has shared the file with you.');
    return;
  }

  // 3. Check for duplicates
  const ssKey = 'Codex';
  const sheetName = 'Custom Abilities';
  const codexSS = fGetCodexSpreadsheet();
  const { arr, rowTags, colTags } = fGetSheetData(ssKey, sheetName, codexSS, true);
  const headerRow = rowTags.header;
  const sheetIdCol = colTags.sheetid;

  for (let r = headerRow + 1; r < arr.length; r++) {
    if (arr[r][sheetIdCol] === sourceId) {
      fEndToast();
      fShowMessage('⚠️ Duplicate', 'This custom source has already been added to your Codex.');
      return;
    }
  }

  // 4. Prompt for a friendly name with the updated example text
  const sourceName = fPromptWithInput('Name the Source', `✅ Success! File access verified.\n\nOwner: ${ownerEmail}\n\nPlease enter a friendly name for this source (e.g., "John's Custom List"):`);
  if (!sourceName) {
    fEndToast();
    fShowMessage('ℹ️ Canceled', 'Operation was canceled.');
    return;
  }

  // 5. Add the new source to the sheet
  fShowToast('Adding new source to your Codex...', 'Add New Source');
  const destSheet = codexSS.getSheetByName(sheetName);
  const lastRow = destSheet.getLastRow();
  let targetRow;

  const dataToWrite = [];
  dataToWrite[colTags.sheetid - 1] = sourceId;
  dataToWrite[colTags.custabilitiesname - 1] = sourceName; // <-- CHANGE HERE
  dataToWrite[colTags.owner - 1] = (ownerEmail === currentUser) ? 'Me' : ownerEmail; // <-- CHANGE HERE

  const firstDataRowIndex = headerRow + 1;
  const templateRow = firstDataRowIndex + 1; // 1-based template row
  const nameCol = colTags.custabilitiesname; // <-- CHANGE HERE

  if (arr.length <= firstDataRowIndex || !arr[firstDataRowIndex][nameCol]) {
    targetRow = templateRow;
  } else {
    targetRow = lastRow + 1;
    destSheet.insertRowsAfter(lastRow, 1);
    const formatSourceRange = destSheet.getRange(templateRow, 1, 1, destSheet.getMaxColumns());
    const formatDestRange = destSheet.getRange(targetRow, 1, 1, destSheet.getMaxColumns());
    formatSourceRange.copyTo(formatDestRange, { formatOnly: true });
  }

  const targetRange = destSheet.getRange(targetRow, 2, 1, dataToWrite.length);
  targetRange.setValues([dataToWrite]);

  fEndToast();
  fShowMessage('✅ Success', `The custom source "${sourceName}" has been successfully added to your Codex.`);
} // End function fAddNewCustomSource

/* function fAddOwnCustomAbilitiesSource
   Purpose: Automatically finds the player's own 'Cust' file and logs it as the first entry in <Custom Abilities>.
   Assumptions: This is run at the end of the initial setup, so the <MyVersions> sheet is populated.
   Notes: Ensures the player always has access to their own custom content.
   @returns {void}
*/
function fAddOwnCustomAbilitiesSource() {
  const codexSS = fGetCodexSpreadsheet();

  // 1. Find the player's 'Cust' file ID from their local <MyVersions> sheet.
  const custId = fGetSheetId(g.CURRENT_VERSION, 'Cust');

  // 2. Get the destination sheet and its properties.
  const destSheetName = 'Custom Abilities';
  const destSS = SpreadsheetApp.getActiveSpreadsheet();
  const destSheet = destSS.getSheetByName(destSheetName);
  const { arr, rowTags, colTags } = fGetSheetData('Codex', destSheetName, codexSS, true);
  const headerRow = rowTags.header;
  const firstDataRow = headerRow + 2; // 1-based row number for the first data entry

  // 3. Prepare the data to be written.
  const dataToWrite = [];
  dataToWrite[colTags.sheetid - 1] = custId;
  dataToWrite[colTags.custabilitiesname - 1] = 'My Custom Abilities'; // <-- CHANGE HERE
  dataToWrite[colTags.owner - 1] = 'Me';

  // 4. Write the data to the first pre-formatted row.
  const targetRange = destSheet.getRange(firstDataRow, 2, 1, dataToWrite.length);
  targetRange.setValues([dataToWrite]);
} // End function fAddOwnCustomAbilitiesSource