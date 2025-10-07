/* global SpreadsheetApp */
/* exported fVerCreateMenu */

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
  const createMenu = SpreadsheetApp.getUi().createMenu('Create New Character')
    .addItem('Latest Version', 'fMenuCreateLatestCharacter')
    .addItem('Older Legacy Version', 'fMenuCreateLegacyCharacter');

  SpreadsheetApp.getUi()
    .createMenu('*** Flex ***')
    .addSubMenu(createMenu)
    .addItem('Rename Character', 'fMenuRenameCharacter')
    .addSeparator()
    .addItem('Delete Character(s)', 'fMenuDeleteCharacter')
    .addToUi();
} // End function fCreateCodexMenu

/* function fCreateFlexMenu
   Purpose: Creates the main custom menu for Flex spreadsheets.
   Assumptions: This is called from an onOpen trigger.
   Notes: This will be the primary user-facing menu. Includes a placeholder item to prevent errors.
   @returns {void}
*/
function fCreateFlexMenu() {
  SpreadsheetApp.getUi()
    .createMenu('*** Flex ***')
    .addItem('Filter Powers⚡', 'fMenuFilterPowers')
    .addItem('More Actions Coming Soon...', 'fMenuPlaceholder')
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
  menu.addItem('Show/Hide All', 'fMenuToggleVisibility');
  menu.addSeparator();

  // Context-specific items
  if (context === 'DB') {
    menu.addItem('Build Powers', 'fMenuBuildPowers');
  }
  if (context === 'CS') {
    menu.addItem('Update <Choose Powers>', 'fMenuUpdatePowerTables');
    menu.addItem('Copy CS <Game> to <Paper>', 'fMenuPrepGameForPaper'); // <-- RENAMED
  }

  menu.addSeparator();
  menu.addItem('Test', 'fMenuTest');
  menu.addToUi();
} // End function fCreateDesignerMenu