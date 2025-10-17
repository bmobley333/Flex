/* global SpreadsheetApp */
/* exported fCreateCodexMenu, fCreateFlexMenu, fCreateGenericMenus, fCreateCustMenu, fCreateDesignerMenu */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Menu Creation
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


/* function fCreateCodexMenu
   Purpose: Creates the main custom menu for the Codex spreadsheet.
   Assumptions: This is called from the onOpen trigger of the Codex sheet.
   Notes: This will be the primary user-facing menu.
   @returns {void}
*/
function fCreateCodexMenu() {
  const ui = SpreadsheetApp.getUi();

  // --- Characters Submenu ---
  const charactersMenu = ui.createMenu('👤 Characters')
    .addItem('Create New', 'fMenuCreateLatestCharacter')
    .addItem('Create Old Legacy Version', 'fMenuCreateLegacyCharacter')
    .addItem('Rename', 'fMenuRenameCharacter')
    .addSeparator()
    .addItem('Delete Character(s)', 'fMenuDeleteCharacter');

  // --- Custom Abilities Submenu ---
  const customAbilitiesMenu = ui.createMenu('⚡ Custom Abilities')
    .addItem('Create New Sheet', 'fMenuCreateCustomList')
    .addItem('Rename Sheet', 'fMenuRenameCustomList')
    .addItem('Delete Sheet(s)', 'fMenuDeleteCustomList')
    .addSeparator()
    .addItem('Share My Sheet(s)', 'fMenuShareCustomLists')
    .addItem('Add Sheet From ID', 'fMenuAddNewCustomSource');

  // --- Main Flex Menu ---
  ui.createMenu('💪 MS3')
    .addSubMenu(charactersMenu)
    .addSubMenu(customAbilitiesMenu)
    .addToUi();
} // End function fCreateCodexMenu

/* function fCreateFlexMenu
   Purpose: Creates the main custom menu for Flex spreadsheets.
   Assumptions: This is called from an onOpen trigger.
   Notes: This will be the primary user-facing menu.
   @returns {void}
*/
function fCreateFlexMenu() {
  const ui = SpreadsheetApp.getUi();
  const filterPowersMenu = ui.createMenu('⚡ Filter Powers')
    .addItem('Load All DB and Cust Powers', 'fMenuSyncPowerChoices')
    .addItem('Filter Powers From Selections ⚡', 'fMenuFilterPowers')
    .addSeparator()
    .addItem('Clear All Selections', 'fMenuClearPowerChoices');

  const filterMagicItemsMenu = ui.createMenu('✨ Filter Magic Items')
    .addItem('Load All DB and Cust Items', 'fMenuSyncMagicItemChoices')
    .addItem('Filter Items From Selections ✨', 'fMenuFilterMagicItems')
    .addSeparator()
    .addItem('Clear All Selections', 'fMenuClearMagicItemChoices');

  const filterSkillSetsMenu = ui.createMenu('🎓 Filter Skill Sets')
    .addItem('Load All DB and Cust Skill Sets', 'fMenuSyncSkillSetChoices')
    .addItem('Filter Skill Sets From Selections 🎓', 'fMenuFilterSkillSets')
    .addSeparator()
    .addItem('Clear All Selections', 'fMenuClearSkillSetChoices');

  ui.createMenu('💪 MS3')
    .addSubMenu(filterPowersMenu)
    .addSubMenu(filterMagicItemsMenu)
    .addSubMenu(filterSkillSetsMenu)
    .addToUi();
} // End function fCreateFlexMenu


/* function fCreateGenericMenus
   Purpose: Creates the standard set of menus for most sheets.
   Assumptions: This is called from an onOpen trigger.
   Notes: A wrapper function to ensure both the Flex and Designer menus are created.
   @param {string} context - The context of the sheet (e.g., 'CS', 'DB').
   @returns {void}
*/
function fCreateGenericMenus(context) {
  fCreateFlexMenu();
  fCreateDesignerMenu(context);
} // End function fCreateGenericMenus

/* function fCreateCustMenu
   Purpose: Creates the main custom menu for the Custom Abilities spreadsheet.
   Assumptions: This is called from an onOpen trigger in a Cust sheet.
   Notes: This menu provides tools for managing powers and items within the sheet.
   @returns {void}
*/
function fCreateCustMenu() {
  const ui = SpreadsheetApp.getUi();

  const powersMenu = ui.createMenu('⚡ Powers')
    .addItem('✅ Verify & Publish Powers', 'fMenuVerifyAndPublish')
    .addSeparator()
    .addItem('🗑️ Delete Selected Powers', 'fMenuDeleteSelectedPowers');

  const magicItemsMenu = ui.createMenu('✨ Magic Items')
    .addItem('✅ Verify & Publish Items', 'fMenuVerifyAndPublishMagicItems')
    .addSeparator()
    .addItem('🗑️ Delete Selected Items', 'fMenuDeleteSelectedMagicItems');

  const skillSetsMenu = ui.createMenu('🎓 Skill Sets')
    .addItem('✅ Verify & Publish Skill Sets', 'fMenuVerifyAndPublishSkillSets')
    .addSeparator()
    .addItem('🗑️ Delete Selected Skill Sets', 'fMenuDeleteSelectedSkillSets');

  ui.createMenu('💪 MS3')
    .addSubMenu(powersMenu)
    .addSubMenu(magicItemsMenu)
    .addSubMenu(skillSetsMenu)
    .addToUi();
} // End function fCreateCustMenu


/* function fCreateDesignerMenu
   Purpose: Creates the generic "Designer" custom menu, customized by context.
   Assumptions: This is called from an onOpen trigger.
   Notes: This can be used by any sheet to create a consistent designer menu.
   @param {string} [context=''] - The context of the sheet ('CS', 'DB', 'Codex', etc.).
   @returns {void}
*/
function fCreateDesignerMenu(context = '') {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('⚙️Designer');

  // Context-specific items
  if (context === 'DB') {
    const powersSubMenu = ui.createMenu('⚡ Powers')
      .addItem('Build Powers from Tables', 'fMenuBuildPowers');
    const magicItemsSubMenu = ui.createMenu('✨ Magic Items')
      .addItem('Build Magic Items from Tables', 'fMenuBuildMagicItems');
    const skillSetsSubMenu = ui.createMenu('🎓 Skill Sets')
      .addItem('Build Skill Sets from Tables', 'fMenuBuildSkillSets');
    menu.addSubMenu(powersSubMenu);
    menu.addSubMenu(magicItemsSubMenu);
    menu.addSubMenu(skillSetsSubMenu);
    menu.addSeparator();
  }

  if (context === 'CS') {
    menu.addItem('Copy CS <Game> to <Paper>', 'fMenuPrepGameForPaper');
    menu.addSeparator();
  }

  if (context === 'Tables') {
    const skillsSubMenu = ui.createMenu('🎓 Skills')
      .addItem('Verify Skill Types', 'fMenuVerifyIndividualSkills');
    const skillSetsSubMenu = ui.createMenu('🎓 Skill Sets')
      .addItem('Verify', 'fMenuVerifySkillSetLists');
    menu.addSubMenu(skillsSubMenu);
    menu.addSubMenu(skillSetsSubMenu);
    menu.addSeparator();
  }


  menu.addItem('Tag Verification', 'fMenuTagVerification');
  menu.addItem('Trim Empty Rows/Cols', 'fMenuTrimSheet');
  menu.addItem('Show/Hide All', 'fMenuToggleVisibility');
  menu.addSeparator();
  menu.addItem('Test', 'fMenuTest');
  menu.addToUi();
} // End function fCreateDesignerMenu