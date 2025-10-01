/* global SpreadsheetApp */
/* exported fVerCreateMenu */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Menu Creation
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

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
    .addSeparator()
    .addItem('Show/Hide All', 'fMenuToggleVisibility')
    .addToUi();
} // End function fCreateDesignerMenu