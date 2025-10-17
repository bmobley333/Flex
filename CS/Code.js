/* global FlexLib, SpreadsheetApp */

// --- Session Caches for High-Speed Performance ---
let powerDataCache = null; // Caches the filtered power data.
let magicItemDataCache = null; // Caches the filtered magic item data.

const SCRIPT_INITIALIZED_KEY = 'SCRIPT_INITIALIZED';


/* function onOpen
   Purpose: Simple trigger that runs automatically when the spreadsheet is opened.
   Assumptions: None.
   Notes: Builds menus based on authorization status and user identity (player vs. designer).
   @returns {void}
*/
function onOpen() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const isInitialized = scriptProperties.getProperty(SCRIPT_INITIALIZED_KEY);

  if (isInitialized) {
    // Always create the main player menu.
    FlexLib.fCreateFlexMenu();

    // Get the globals object from the library.
    const g = FlexLib.getGlobals();

    // Only show the Designer menu if the user is the admin.
    if (Session.getActiveUser().getEmail() === g.ADMIN_EMAIL) {
      FlexLib.fCreateDesignerMenu('CS');
    }
  } else {
    SpreadsheetApp.getUi()
      .createMenu('💪 MS3')
      .addItem('▶️ Activate 💪MS3 Menus', 'fActivateMenus')
      .addToUi();
  }
} // End function onOpen

/* function fActivateMenus
   Purpose: Runs the first-time authorization and menu setup.
   Assumptions: Triggered by a user clicking the 'Activate' menu item.
   Notes: This function's execution by a user triggers the Google Auth prompt if needed.
   @returns {void}
*/
function fActivateMenus() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty(SCRIPT_INITIALIZED_KEY, 'true');

  // --- NEW ---
  // Run the one-time character sheet onboarding process.
  FlexLib.run('CharacterOnboarding');
  // --- END NEW ---

  const title = 'IMPORTANT - Please Refresh Browser Tab';
  const message = '✅ Success! The script has been authorized and your sheet has been set up with all core game choices.\n\nPlease refresh this browser tab now to load the full custom menus.';
  SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
} // End function fActivateMenus



/* function fGetSkillsFromString
   Purpose: Parses a comma-separated string of skills into a clean array.
   Assumptions: None.
   Notes: A helper for fProcessSkillSetChange.
   @param {string} skillString - The raw CSV string of skills.
   @returns {string[]} An array of cleaned, individual skill strings.
*/
function fGetSkillsFromString(skillString) {
  if (!skillString) return [];
  return skillString
    .split(',')
    .map(s => s.trim())
    .filter(s => s); // Remove empty strings
} // End function fGetSkillsFromString

/* function fUpdateCharacterSkills
   Purpose: Adds or removes a list of skills from the appropriate sections on the <Game> sheet.
   Assumptions: None.
   Notes: A helper for fProcessSkillSetChange that contains the core placement and removal logic.
   @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The <Game> sheet object.
   @param {string[]} skills - An array of skill strings to process.
   @param {object} gameColTags - The column tags for the <Game> sheet.
   @param {string} mode - The operation mode, either 'ADD' or 'REMOVE'.
   @returns {void}
*/
function fUpdateCharacterSkills(sheet, skills, gameColTags, mode) {
  const emojiMap = { '💪': 'mightskills', '🏃': 'motionskills', '👁️': 'mindskills', '✨': 'magicskills' };
  const validEmojis = Object.keys(emojiMap);
  const individualSkillsCol = gameColTags.individualskills + 1;

  skills.forEach(skillWithEmoji => {
    let detectedEmoji = null;
    for (const emoji of validEmojis) {
      if (skillWithEmoji.endsWith(emoji)) {
        detectedEmoji = emoji;
        break;
      }
    }

    if (!detectedEmoji) {
      FlexLib.fShowMessage('⚠️ Invalid Skill', `The skill "${skillWithEmoji}" has an invalid type and was skipped.`);
      return;
    }

    const targetRowTag = emojiMap[detectedEmoji];
    const { rowTags: gameRowTags } = FlexLib.fGetSheetData('CS', 'Game');
    const baseRowIndex = gameRowTags[targetRowTag] + 1;

    // The full skill string (e.g., "Infrared👁️") is now the identifier.
    const skillIdentifier = skillWithEmoji;

    if (mode === 'ADD') {
      const row1Range = sheet.getRange(baseRowIndex, individualSkillsCol);
      const row2Range = sheet.getRange(baseRowIndex + 1, individualSkillsCol);
      const row1Text = row1Range.getValue();
      const row2Text = row2Range.getValue();

      // Find if the skill already exists in either row to increment its count.
      let foundInRow = null;
      let existingSkills = [];

      // Check Row 1
      existingSkills = row1Text ? row1Text.split(',').map(s => s.trim()) : [];
      let skillIndex = existingSkills.findIndex(s => s.endsWith(skillIdentifier));
      if (skillIndex !== -1) {
        foundInRow = { range: row1Range, skills: existingSkills, index: skillIndex };
      } else {
        // Check Row 2
        existingSkills = row2Text ? row2Text.split(',').map(s => s.trim()) : [];
        skillIndex = existingSkills.findIndex(s => s.endsWith(skillIdentifier));
        if (skillIndex !== -1) {
          foundInRow = { range: row2Range, skills: existingSkills, index: skillIndex };
        }
      }

      if (foundInRow) {
        // --- Skill exists, increment the count ---
        const existingSkill = foundInRow.skills[foundInRow.index];
        const parts = existingSkill.split('_');
        const count = parts.length > 1 ? parseInt(parts[0], 10) + 1 : 2;
        foundInRow.skills[foundInRow.index] = `${count}_${skillIdentifier}`;
        foundInRow.range.setValue(foundInRow.skills.join(', '));
      } else {
        // --- Skill is new, add it to the shorter row ---
        const targetRange = row1Text.length <= row2Text.length ? row1Range : row2Range;
        const currentText = targetRange.getValue();
        const newText = currentText ? `${currentText}, ${skillIdentifier}` : skillIdentifier;
        targetRange.setValue(newText);
      }
    } else if (mode === 'REMOVE') {
      // --- Decrement or remove the skill ---
      for (let i = 0; i < 2; i++) {
        const range = sheet.getRange(baseRowIndex + i, individualSkillsCol);
        const text = range.getValue();
        if (!text) continue;

        const existingSkills = text.split(',').map(s => s.trim());
        const skillIndex = existingSkills.findIndex(s => s.endsWith(skillIdentifier));

        if (skillIndex !== -1) {
          const existingSkill = existingSkills[skillIndex];
          const parts = existingSkill.split('_');
          if (parts.length > 1) {
            const count = parseInt(parts[0], 10) - 1;
            if (count > 1) {
              existingSkills[skillIndex] = `${count}_${skillIdentifier}`;
            } else {
              existingSkills[skillIndex] = skillIdentifier;
            }
          } else {
            existingSkills.splice(skillIndex, 1);
          }
          range.setValue(existingSkills.join(', '));
          break; // Exit after processing
        }
      }
    }
  });
} // End function fUpdateCharacterSkills

/* function onEdit
   Purpose: A simple trigger that auto-populates details from a high-speed session cache when an item is selected from a dropdown.
   Assumptions: The appropriate DataCache sheet exists. The <Game> sheet is tagged correctly.
   Notes: This is the optimized auto-formatter, built on fGetSheetData for maximum performance and robust, explicit tag matching.
   @param {GoogleAppsScript.Events.SheetsOnEdit} e - The event object passed by the trigger.
   @returns {void}
*/
function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== 'Game') return;

  try {
    const { colTags: gameColTags } = FlexLib.fGetSheetData('CS', 'Game', e.source);
    const editedColTag = Object.keys(gameColTags).find(tag => gameColTags[tag] === e.range.getColumn() - 1);

    if (!editedColTag) return;

    // 1. Build high-speed maps from the data caches
    const { arr: powerArr, colTags: powerColTags } = FlexLib.fGetSheetData('CS', 'PowerDataCache', e.source);
    const powerMap = new Map();
    powerArr.slice(1).forEach(row => {
      const dropdownText = row[powerColTags.dropdown];
      if (dropdownText) {
        powerMap.set(dropdownText, {
          usage: row[powerColTags.usage],
          action: row[powerColTags.action],
          name: row[powerColTags.abilityname],
          effect: row[powerColTags.effect],
        });
      }
    });

    const { arr: itemArr, colTags: itemColTags } = FlexLib.fGetSheetData('CS', 'MagicItemDataCache', e.source);
    const magicItemMap = new Map();
    itemArr.slice(1).forEach(row => {
      const dropdownText = row[itemColTags.dropdown];
      if (dropdownText) {
        magicItemMap.set(dropdownText, {
          usage: row[itemColTags.usage],
          action: row[itemColTags.action],
          name: row[itemColTags.abilityname],
          effect: row[itemColTags.effect],
        });
      }
    });

    const { arr: skillSetArr, colTags: skillSetColTags } = FlexLib.fGetSheetData('CS', 'SkillSetDataCache', e.source);
    const skillSetMap = new Map();
    skillSetArr.slice(1).forEach(row => {
      const dropdownText = row[skillSetColTags.dropdown];
      if (dropdownText) {
        skillSetMap.set(dropdownText, { name: row[skillSetColTags.name], effect: row[skillSetColTags.effect] });
      }
    });

    // 2. EXPLICIT TAG MAPPING & ACTION
    switch (editedColTag) {
      case 'powerdropdown1':
      case 'magicitemdropdown1':
      case 'powerdropdown2':
      case 'magicitemdropdown2': {
        const selectedValue = e.value;
        const data = powerMap.has(selectedValue) ? powerMap.get(selectedValue) : magicItemMap.get(selectedValue);
        const isPower = powerMap.has(selectedValue);
        const isDropdown1 = editedColTag.endsWith('1');

        const pUsage = isDropdown1 ? 'powerusage1' : 'powerusage2';
        const pAction = isDropdown1 ? 'poweraction1' : 'poweraction2';
        const pName = isDropdown1 ? 'powername1' : 'powername2';
        const pEffect = isDropdown1 ? 'powereffect1' : 'powereffect2';

        const mUsage = isDropdown1 ? 'magicitemusage1' : 'magicitemusage2';
        const mAction = isDropdown1 ? 'magicitemaction1' : 'magicitemaction2';
        const mName = isDropdown1 ? 'magicitemname1' : 'magicitemname2';
        const mEffect = isDropdown1 ? 'magicitemeffect1' : 'magicitemeffect2';

        [pUsage, pAction, pName, pEffect, mUsage, mAction, mName, mEffect].forEach(tag => {
          const col = gameColTags[tag];
          if (col !== undefined) sheet.getRange(e.range.getRow(), col + 1).clearContent();
        });

        if (data) {
          const usageCol = gameColTags[isPower ? pUsage : mUsage];
          const actionCol = gameColTags[isPower ? pAction : mAction];
          const nameCol = gameColTags[isPower ? pName : mName];
          const effectCol = gameColTags[isPower ? pEffect : mEffect];

          if (usageCol !== undefined) sheet.getRange(e.range.getRow(), usageCol + 1).setValue(data.usage);
          if (actionCol !== undefined) sheet.getRange(e.range.getRow(), actionCol + 1).setValue(data.action);
          if (nameCol !== undefined) sheet.getRange(e.range.getRow(), nameCol + 1).setValue(data.name);
          if (effectCol !== undefined) sheet.getRange(e.range.getRow(), effectCol + 1).setValue(data.effect);
        }
        break;
      }

      case 'skillsetdropdown': {
        const nameCol = gameColTags.skillsetname;
        const effectCol = gameColTags.skillseteffect;
        const editedRow = e.range.getRow();

        // --- REMOVAL LOGIC ---
        // Must run first, using the skill list currently on the sheet before it gets cleared.
        if (e.oldValue) {
          const effectString = sheet.getRange(editedRow, effectCol + 1).getValue();
          const skillsToRemove = fGetSkillsFromString(effectString);
          if (skillsToRemove.length > 0) {
            FlexLib.fShowToast('⏳ Removing old skill set...', 'Skill Sets');
            fUpdateCharacterSkills(sheet, skillsToRemove, gameColTags, 'REMOVE');
          }
        }

        // --- CLEARING / ADDITION LOGIC ---
        if (!e.value) { // Cell was cleared
          if (nameCol !== undefined) sheet.getRange(editedRow, nameCol + 1).clearContent();
          if (effectCol !== undefined) sheet.getRange(editedRow, effectCol + 1).clearContent();
        } else { // New value was added
          const data = skillSetMap.get(e.value);
          if (data) {
            if (nameCol !== undefined) sheet.getRange(editedRow, nameCol + 1).setValue(data.name);
            if (effectCol !== undefined) sheet.getRange(editedRow, effectCol + 1).setValue(data.effect);
            const skillsToAdd = fGetSkillsFromString(data.effect);
            if (skillsToAdd.length > 0) {
              FlexLib.fShowToast('⏳ Adding new skill set...', 'Skill Sets');
              fUpdateCharacterSkills(sheet, skillsToAdd, gameColTags, 'ADD');
            }
          }
        }
        
        if (e.oldValue || e.value) FlexLib.fEndToast();
        break;
      }

      default:
        return;
    }
  } catch (err) {
    console.error(`❌ CRITICAL ERROR in onEdit: ${err.message}\n${err.stack}`);
  }
} // End function onEdit



/* function buttonFilterPowers
   Purpose: Local trigger for a button, mimics the "Filter Powers" menu item.
   Assumptions: None.
   Notes: Assign this function name to a button in the sheet to trigger the FilterPowers command.
   @returns {void}
*/
function buttonFilterPowers() {
  FlexLib.run('FilterPowers');
} // End function buttonFilterPowers


/* function buttonFilterMagicItems
   Purpose: Local trigger for a button, mimics the "Filter Magic Items" menu item.
   Assumptions: None.
   Notes: Assign this function name to a button in the sheet to trigger the FilterMagicItems command.
   @returns {void}
*/
function buttonFilterMagicItems() {
  FlexLib.run('FilterMagicItems');
} // End function buttonFilterMagicItems

/* function buttonClearPowerChoices
   Purpose: Local trigger for a button to clear all power filter checkboxes.
   Assumptions: None.
   Notes: Assign this function name to a button on the <Filter Powers> sheet.
   @returns {void}
*/
function buttonClearPowerChoices() {
  FlexLib.run('ClearPowerFilters');
} // End function buttonClearPowerChoices


/* function buttonClearMagicItemChoices
   Purpose: Local trigger for a button to clear all magic item filter checkboxes.
   Assumptions: None.
   Notes: Assign this function name to a button on the <Filter Magic Items> sheet.
   @returns {void}
*/
function buttonClearMagicItemChoices() {
  FlexLib.run('ClearMagicItemFilters');
} // End function buttonClearMagicItemChoices

/* function onChange
   Purpose: An installable trigger that invalidates the session cache for the <Game> sheet when its structure changes.
   Assumptions: This trigger is manually installed for the spreadsheet.
   Notes: This protects against data corruption if a user inserts/deletes rows or columns.
   @param {GoogleAppsScript.Events.SheetsOnChange} e - The event object passed by the trigger.
   @returns {void}
*/
function onChange(e) {
  // --- THIS IS THE FIX ---
  // We only care about structural changes on the Game sheet.
  if (e.source.getActiveSheet().getName() !== 'Game') return;

  const structuralChanges = ['INSERT_ROW', 'REMOVE_ROW', 'INSERT_COLUMN', 'REMOVE_COLUMN'];
  if (structuralChanges.includes(e.changeType)) {
    // A structural change was made, so we call the library to invalidate the central cache.
    FlexLib.run('InvalidateGameCache');
  }
} // End function onChange


/* function fMenuTagVerification
   Purpose: The local trigger function called by the "Tag Verification" menu item.
   Assumptions: None.
   Notes: This function acts as a simple pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuTagVerification() {
  FlexLib.run('TagVerification');
} // End function fMenuTagVerification

/* function fMenuToggleVisibility
   Purpose: Local trigger for the "Show/Hide All" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuToggleVisibility() {
  FlexLib.run('ToggleVisibility');
} // End function fMenuToggleVisibility

/* function fMenuTest
   Purpose: Local trigger for the "Test" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuTest() {
  FlexLib.run('Test');
} // End function fMenuTest

/* function fMenuSyncPowerChoices
   Purpose: Local trigger for the "Sync Power Choices" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuSyncPowerChoices() {
  FlexLib.run('SyncPowerChoices');
} // End function fMenuSyncPowerChoices

/* function fMenuFilterPowers
   Purpose: Local trigger for the "Filter Powers" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuFilterPowers() {
  FlexLib.run('FilterPowers');
} // End function fMenuFilterPowers

/* function fMenuSyncMagicItemChoices
   Purpose: Local trigger for the "Sync Magic Item Choices" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuSyncMagicItemChoices() {
  FlexLib.run('SyncMagicItemChoices');
} // End function fMenuSyncMagicItemChoices

/* function fMenuFilterMagicItems
   Purpose: Local trigger for the "Filter Magic Items" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuFilterMagicItems() {
  FlexLib.run('FilterMagicItems');
} // End function fMenuFilterMagicItems

/* function fMenuPrepGameForPaper
   Purpose: Local trigger for the "Copy CS <Game> to <Paper>" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuPrepGameForPaper() {
  FlexLib.run('PrepGameForPaper');
} // End function fMenuPrepGameForPaper

/* function fMenuClearPowerChoices
   Purpose: Local trigger for the "Clear All Selections" menu item for powers.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuClearPowerChoices() {
  FlexLib.run('ClearPowerFilters');
} // End function fMenuClearPowerChoices

/* function fMenuClearMagicItemChoices
   Purpose: Local trigger for the "Clear All Selections" menu item for magic items.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuClearMagicItemChoices() {
  FlexLib.run('ClearMagicItemFilters');
} // End function fMenuClearMagicItemChoices

/* function buttonFilterSkillSets
   Purpose: Local trigger for a button, mimics the "Filter Skill Sets" menu item.
   Assumptions: None.
   Notes: Assign this function name to a button in the sheet to trigger the FilterSkillSets command.
   @returns {void}
*/
function buttonFilterSkillSets() {
  FlexLib.run('FilterSkillSets');
} // End function buttonFilterSkillSets

/* function buttonClearSkillSetChoices
   Purpose: Local trigger for a button to clear all skill set filter checkboxes.
   Assumptions: None.
   Notes: Assign this function name to a button on the <Filter Skill Sets> sheet.
   @returns {void}
*/
function buttonClearSkillSetChoices() {
  FlexLib.run('ClearSkillSetFilters');
} // End function buttonClearSkillSetChoices

/* function fMenuSyncSkillSetChoices
   Purpose: Local trigger for the "Sync Skill Set Choices" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuSyncSkillSetChoices() {
  FlexLib.run('SyncSkillSetChoices');
} // End function fMenuSyncSkillSetChoices

/* function fMenuFilterSkillSets
   Purpose: Local trigger for the "Filter Skill Sets" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuFilterSkillSets() {
  FlexLib.run('FilterSkillSets');
} // End function fMenuFilterSkillSets

/* function fMenuClearSkillSetChoices
   Purpose: Local trigger for the "Clear All Selections" menu item for skill sets.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuClearSkillSetChoices() {
  FlexLib.run('ClearSkillSetFilters');
} // End function fMenuClearSkillSetChoices