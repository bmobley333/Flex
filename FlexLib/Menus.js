/* global SpreadsheetApp */
/* exported fVerCreateMenu */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Menu Creation
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fVerCreateMenu
   Purpose: Creates the custom menu for the Ver (Version Tracker) sheet.
   Assumptions: This is called from the onOpen trigger of the Ver sheet.
   Notes: All menu logic for the Ver sheet will be managed here.
   @returns {void}
*/
function fVerCreateMenu() {
  SpreadsheetApp.getUi()
    .createMenu('Designer')
    .addItem('Tag Verification', 'fVerMenuTagVerification')
    .addToUi();
} // End function fVerCreateMenu