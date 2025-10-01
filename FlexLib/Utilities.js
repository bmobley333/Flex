/* global SpreadsheetApp */
/* exported fShowMessage */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - User Interface Utilities
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fShowMessage
   Purpose: Displays a simple modal pop-up message to the user.
   Assumptions: None.
   Notes: This will be our standard method for all user-facing modal alerts.
   @param {string} title - The title to display in the message box header.
   @param {string} message - The main body of the message.
   @returns {void}
*/
function fShowMessage(title, message) {
  SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
} // End function fShowMessage