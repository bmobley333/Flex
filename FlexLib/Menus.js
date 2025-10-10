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
  const charactersMenu = ui.createMenu('Characters')
    .addItem('Create New', 'fMenuCreateLatestCharacter')
    .addItem('Create Old Legacy Version', 'fMenuCreateLegacyCharacter')
    .addItem('Rename', 'fMenuRenameCharacter')
    .addSeparator()
    .addItem('Delete Character(s)', 'fMenuDeleteCharacter');

  // --- Custom Abilities Submenu ---
  const customAbilitiesMenu = ui.createMenu('Custom Abilities')
    .addItem('Create New Sheet', 'fMenuCreateCustomList')
    .addItem('Rename Sheet', 'fMenuRenameCustomList')
    .addItem('Delete Sheet(s)', 'fMenuDeleteCustomList')
    .addSeparator()
    .addItem('Share My Sheet(s)', 'fMenuShareCustomLists')
    .addItem('Add Sheet From ID', 'fMenuAddNewCustomSource');

  // --- Main Flex Menu ---
  ui.createMenu('💪 Flex')
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
  const filterMenu = SpreadsheetApp.getUi().createMenu('Filter Powers')
    .addItem('Load All DB and Cust Powers', 'fMenuSyncPowerChoices')
    .addItem('Filter Powers From Selections ⚡', 'fMenuFilterPowers');

  SpreadsheetApp.getUi()
    .createMenu('💪 Flex')
    .addSubMenu(filterMenu)
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
   Notes: This menu provides tools for managing powers within the sheet.
   @returns {void}
*/
function fCreateCustMenu() {
  const powersMenu = SpreadsheetApp.getUi().createMenu('⚡ Powers')
    .addItem('✅ Verify & Publish Powers', 'fMenuVerifyAndPublish')
    .addSeparator()
    .addItem('🗑️ Delete Selected Powers', 'fMenuDeleteSelectedPowers');

  SpreadsheetApp.getUi()
    .createMenu('💪 Flex')
    .addSubMenu(powersMenu)
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
  const menu = SpreadsheetApp.getUi().createMenu('Designer');

  menu.addItem('Tag Verification', 'fMenuTagVerification');
  menu.addItem('Trim Empty Rows/Cols', 'fMenuTrimSheet');
  menu.addItem('Show/Hide All', 'fMenuToggleVisibility');
  menu.addSeparator();

  // Context-specific items
  if (context === 'DB') {
    menu.addItem('Build Powers', 'fMenuBuildPowers');
  }
  if (context === 'CS') {
    menu.addItem('Copy CS <Game> to <Paper>', 'fMenuPrepGameForPaper');
  }

  menu.addSeparator();
  menu.addItem('Test', 'fMenuTest');
  menu.addToUi();
} // End function fCreateDesignerMenu