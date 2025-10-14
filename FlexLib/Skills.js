/* global fShowToast, SpreadsheetApp, fGetSheetData, fShowMessage, fEndToast, fPromptWithInput */
/* exported fVerifySkills */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Skill Verification
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fVerifySkills
   Purpose: The master workflow for verifying the skill type emoji in the active sheet.
   Assumptions: Run from a 'Tables' sheet context. The active sheet has a 'Header' row tag and a 'skills' column tag.
   Notes: Iterates through all data rows and uses a helper to validate and correct each skill string.
   @returns {void}
*/
function fVerifySkills() {
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
    console.error(`❌ CRITICAL ERROR in fVerifySkills: ${e.message}\n${e.stack}`);
    fEndToast();
    fShowMessage('❌ Error', `A critical error occurred. Please check the execution logs. Error: ${e.message}`);
  }
} // End function fVerifySkills


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
  const foundEmojis = validEmojis.filter(emoji => skillString.includes(emoji));

  // Case 1: Exactly one valid emoji is found.
  if (foundEmojis.length === 1) {
    const emoji = foundEmojis[0];
    // Auto-correct if the emoji is not at the end of the string.
    if (!skillString.trim().endsWith(emoji)) {
      fShowToast(`Fixing format for: "${skillString}"`, '🎓 Skill Verification', 4);
      const cleanedString = skillString.replace(new RegExp(emoji, 'g'), '').trim();
      return `${cleanedString}${emoji}`;
    }
    // Otherwise, the string is already perfect.
    return skillString;
  }

  // Case 2: Zero or multiple valid emojis are found, requiring user input.
  const choices = validEmojis.map((emoji, index) => `${index + 1}. ${emojiMap[emoji]} ${emoji}`);
  const basePrompt = `The skill "${skillString}" has an invalid type.\n\nPlease choose the correct type to apply:\n\n${choices.join('\n')}\n\nEnter a number from 1 to ${validEmojis.length}.`;
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
      let newString = skillString;
      validEmojis.forEach(emoji => {
        newString = newString.replace(new RegExp(emoji, 'g'), '');
      });
      // Add the correct emoji to the end and trim whitespace.
      return `${newString.trim()}${correctEmoji}`;
    }
    // If input was invalid, the loop will continue and re-prompt.
  }
} // End function fValidateAndCorrectSkillString