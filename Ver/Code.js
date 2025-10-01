/* global FlexLib */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Triggers & Local Functions
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function onOpen
   Purpose: Simple trigger that runs automatically when the spreadsheet is opened.
   Assumptions: None.
   Notes: Its sole job is to call the library to build the custom menu.
   @returns {void}
*/
function onOpen() {
  FlexLib.fVerCreateMenu();
} // End function onOpen

/* function fVerMenuTagVerification
   Purpose: The local trigger function called by the "Tag Verification" menu item.
   Assumptions: None.
   Notes: This function acts as a simple pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fVerMenuTagVerification() {
  FlexLib.run('Ver_TagVerification');
} // End function fVerMenuTagVerification
