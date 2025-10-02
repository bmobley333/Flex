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
  SpreadsheetApp.getUi()
    .createMenu('*** Flex ***')
    .addItem('Create New Character', 'fMenuCreateCharacter')
    .addItem('Delete Character(s)', 'fMenuDeleteCharacter')
    .addToUi();
} // End function fCreateCodexMenu

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