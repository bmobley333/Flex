/* global fShowToast, SpreadsheetApp, fGetSheetData, fShowMessage, fEndToast, fPromptWithInput, g, fGetMasterSheetId, fClearAndWriteData, fActivateSheetByName, fGetVerifiedLocalFile, fGetCodexSpreadsheet, fDeleteTableRow */
/* exported fVerifyIndividualSkills, fVerifySkillSetLists, fBuildSkillSets, fUpdateSkillSetChoices, fFilterSkillSets */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Skill Verification and List Generation
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fGetAllSkillSetTablesList
   Purpose: A helper function to get a definitive, aggregated list of all available skill set tables from the DB.
   Assumptions: None.
   Notes: This is the central source of truth for what skill set tables currently exist.
   @returns {{allSkillSetTables: Array<{tableName: string, source: string}>}} An object containing the aggregated list.
*/
function fGetAllSkillSetTablesList() {
  const dbSkillSetTables = [];

  // 1. Get standard tables from the PLAYER'S LOCAL DB copy.
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

  dbSkillSetTables.sort((a, b) => a.tableName.localeCompare(b.tableName));
  return { allSkillSetTables: dbSkillSetTables };
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
  let dbHeader = [];

  const dbFile = fGetVerifiedLocalFile(g.CURRENT_VERSION, 'DB');
  if (!dbFile) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find or restore your local "DB" file to get skill set data from. Please run initial setup.');
    return null;
  }
  const dbSS = SpreadsheetApp.open(dbFile);
  const { arr: allDbSkillSets, rowTags: dbRowTags, colTags: dbColTags } = fGetSheetData('DB', 'SkillSets', dbSS);
  dbHeader = allDbSkillSets[dbRowTags.header];

  const selectedDbTables = selectedTables.filter(t => t.source === 'DB').map(t => t.tableName);
  if (selectedDbTables.length > 0) {
    const dbSkillSets = allDbSkillSets
      .slice(dbRowTags.header + 1)
      .filter(row => selectedDbTables.includes(row[dbColTags.tablename]));
    allSkillSetsData = allSkillSetsData.concat(dbSkillSets);
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