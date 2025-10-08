/* global FlexLib */

/**
 * Simple trigger that runs automatically when the spreadsheet is opened.
 * Its sole job is to call the library to build the custom menu.
 */
function onOpen() {
  FlexLib.fCreateCustMenu();
} // End function onOpen

/**
 * Local trigger for the "Share My Abilities..." menu item.
 * Acts as a pass-through to the central dispatcher in FlexLib.
 */
function fMenuShareMyAbilities() {
  FlexLib.run('ShareMyAbilities');
} // End function fMenuShareMyAbilities