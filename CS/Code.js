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
      .createMenu('💪 Flex')
      .addItem('▶️ Activate Flex Menus', 'fActivateMenus')
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

/* function fProcessSkillSetChange
   Purpose: Handles the addition or removal of skills on the <Game> sheet when a Skill Set is selected.
   Assumptions: Called by the onEdit trigger.
   Notes: This is the master controller for updating the character's skill lists.
   @param {GoogleAppsScript.Events.SheetsOnEdit} e - The event object from onEdit.
   @param {Map<string, object>} skillSetMap - The map of all available skill sets.
   @param {object} gameColTags - The column tags for the <Game> sheet.
   @returns {void}
*/
function fProcessSkillSetChange(e, skillSetMap, gameColTags) {
  const sheet = e.range.getSheet();
  const oldValue = e.oldValue;
  const newValue = e.value;

  // Case 1: A skill set was removed.
  if (oldValue && skillSetMap.has(oldValue)) {
    FlexLib.fShowToast('⏳ Removing old skill set...', 'Skill Sets');
    const skillsToRemove = fGetSkillsFromString(skillSetMap.get(oldValue).effect);
    fUpdateCharacterSkills(sheet, skillsToRemove, gameColTags, 'REMOVE');
  }

  // Case 2: A new skill set was added.
  if (newValue && skillSetMap.has(newValue)) {
    FlexLib.fShowToast('⏳ Adding new skill set...', 'Skill Sets');
    const skillsToAdd = fGetSkillsFromString(skillSetMap.get(newValue).effect);
    fUpdateCharacterSkills(sheet, skillsToAdd, gameColTags, 'ADD');
  }

  // --- THIS IS THE FIX ---
  // If either an add or remove operation happened, show the completion toast.
  if ((oldValue && skillSetMap.has(oldValue)) || (newValue && skillSetMap.has(newValue))) {
    FlexLib.fEndToast();
  }
} // End function fProcessSkillSetChange

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

    const selectedValue = e.value;
    let data = null;
    let targetTags = {};

    // 1. Use the main, high-performance fGetSheetData cache instead of a manual one.
    const { arr: powerArr, colTags: powerColTags } = FlexLib.fGetSheetData('CS', 'PowerDataCache', e.source);
    const powerMap = new Map();
    powerArr.slice(1).forEach(row => {
      if (row[powerColTags.dropdown]) powerMap.set(row[powerColTags.dropdown], { usage: row[powerColTags.usage], action: row[powerColTags.action], name: row[powerColTags.power], effect: row[powerColTags.effect] });
    });

    const { arr: itemArr, colTags: itemColTags } = FlexLib.fGetSheetData('CS', 'MagicItemDataCache', e.source);
    const magicItemMap = new Map();
    itemArr.slice(1).forEach(row => {
      if (row[itemColTags.dropdown]) magicItemMap.set(row[itemColTags.dropdown], { usage: row[itemColTags.usage], action: row[itemColTags.action], name: row[itemColTags.name], effect: row[itemColTags.effect] });
    });

    const { arr: skillSetArr, colTags: skillSetColTags } = FlexLib.fGetSheetData('CS', 'SkillSetDataCache', e.source);
    const skillSetMap = new Map();
    skillSetArr.slice(1).forEach(row => {
      if (row[skillSetColTags.dropdown]) skillSetMap.set(row[skillSetColTags.dropdown], { name: row[skillSetColTags.name], effect: row[skillSetColTags.effect] });
    });


    // 2. Determine which data to use
    if (powerMap.has(selectedValue)) {
      data = powerMap.get(selectedValue);
    } else if (magicItemMap.has(selectedValue)) {
      data = magicItemMap.get(selectedValue);
    } else if (skillSetMap.has(selectedValue)) {
      data = skillSetMap.get(selectedValue);
    }

    // 3. EXPLICIT TAG MAPPING - No tricky logic
    switch (editedColTag) {
      case 'powerdropdown1':
      case 'magicitemdropdown1':
        targetTags = { usage: 'powerusage1', action: 'poweraction1', name: 'powername1', effect: 'powereffect1', m_usage: 'magicitemusage1', m_action: 'magicitemaction1', m_name: 'magicitemname1', m_effect: 'magicitemeffect1' };
        break;
      case 'powerdropdown2':
      case 'magicitemdropdown2':
        targetTags = { usage: 'powerusage2', action: 'poweraction2', name: 'powername2', effect: 'powereffect2', m_usage: 'magicitemusage2', m_action: 'magicitemaction2', m_name: 'magicitemname2', m_effect: 'magicitemeffect2' };
        break;
      case 'skillsetdropdown':
        targetTags = { name: 'skillsetname' };
        fProcessSkillSetChange(e, skillSetMap, gameColTags); // --- THIS IS THE FIX ---
        break;
      // Add more cases here for DropDown3, DropDown4, etc. if they ever exist
      default:
        return; // Not a dropdown we care about
    }

    // 4. Clear or populate cells
    const allPossibleTags = [targetTags.usage, targetTags.action, targetTags.name, targetTags.effect, targetTags.m_usage, targetTags.m_action, targetTags.m_name, targetTags.m_effect];
    if (!selectedValue || !data) {
      allPossibleTags.forEach(tag => {
        const col = gameColTags[tag];
        if (col !== undefined) sheet.getRange(e.range.getRow(), col + 1).clearContent();
      });
      // Also clear the skill set name cell
      if (gameColTags[targetTags.name] !== undefined) {
        sheet.getRange(e.range.getRow(), gameColTags[targetTags.name] + 1).clearContent();
      }
      return;
    }

    // Determine the correct final set of tags based on the data that was found
    const finalTags = data === powerMap.get(selectedValue)
      ? { usage: targetTags.usage, action: targetTags.action, name: targetTags.name, effect: targetTags.effect }
      : data === magicItemMap.get(selectedValue)
        ? { usage: targetTags.m_usage, action: targetTags.m_action, name: targetTags.m_name, effect: targetTags.m_effect }
        : { name: targetTags.name };


    const usageCol = gameColTags[finalTags.usage];
    const actionCol = gameColTags[finalTags.action];
    const nameCol = gameColTags[finalTags.name];
    const effectCol = gameColTags[finalTags.effect];

    if (usageCol !== undefined) sheet.getRange(e.range.getRow(), usageCol + 1).setValue(data.usage);
    if (actionCol !== undefined) sheet.getRange(e.range.getRow(), actionCol + 1).setValue(data.action);
    if (nameCol !== undefined) sheet.getRange(e.range.getRow(), nameCol + 1).setValue(data.name);
    if (effectCol !== undefined) sheet.getRange(e.range.getRow(), effectCol + 1).setValue(data.effect);

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