/* global SpreadsheetApp */
/* exported fPromptWithInput */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - User Prompts
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fPromptWithInput
   Purpose: Prompts the user for input with a customizable message.
   Assumptions: None.
   Notes: A standardized wrapper for the Ui.prompt method.
   @param {string} title - The title for the prompt dialog.
   @param {string} message - The message to display to the user, often including choices.
   @returns {string|null} The user's text input, or null if they canceled.
*/
function fPromptWithInput(title, message) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() === ui.Button.OK) {
    return response.getResponseText();
  } else {
    return null; // User clicked Cancel or the close button
  }
} // End function fPromptWithInput