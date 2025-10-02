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
    .addItem('More Actions Coming Soon...', 'fMenuPlaceholder')
    .addToUi();
} // End function fCreateFlexMenu


/* function fCreateStandardMenus
   Purpose: Creates the standard set of menus for non-Codex sheets.
   Assumptions: This is called from an onOpen trigger.
   Notes: A wrapper function to ensure both the Flex and Designer menus are created.
   @returns {void}
*/
function fCreateStandardMenus() {
  fCreateFlexMenu();
  fCreateDesignerMenu();
} // End function fCreateStandardMenus

/* function fCreateDesignerMenu
   Purpose: Creates the generic "Designer" custom menu.
   Assumptions: This is called from an onOpen trigger.
   Notes: This can be used by any sheet to create a consistent designer menu.
   @returns {void}
*/
/* function fCreateDesignerMenu
   Purpose: Creates the generic "Designer" custom menu.
   Assumptions: This is called from an onOpen trigger.
   Notes: This can be used by any sheet to create a consistent designer menu.
   @returns {void}
*/
function fCreateDesignerMenu() {
  SpreadsheetApp.getUi()
    .createMenu('Designer')
    .addItem('Tag Verification', 'fMenuTagVerification')
    .addItem('Show/Hide All', 'fMenuToggleVisibility')
    .addSeparator()
    .addItem('Clear PropertiesService', 'fMenuClearProperties')
    .addSeparator()
    .addItem('Test', 'fMenuTest')
    .addToUi();
} // End function fCreateDesignerMenu