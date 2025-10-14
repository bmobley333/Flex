/* global fShowToast, SpreadsheetApp, fGetSheetData, fShowMessage, fEndToast, fPromptWithInput, g, fGetMasterSheetId, fClearAndWriteData, fActivateSheetByName */
/* exported fVerifyIndividualSkills, fVerifySkillSetLists, fBuildSkillSets */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Skill Verification
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


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
      newRow[destColTags.skilllist] = skillList;

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