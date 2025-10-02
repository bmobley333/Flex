/* global FlexLib */

/* function onOpen
   Purpose: Simple trigger that runs automatically when the spreadsheet is opened.
   Assumptions: None.
   Notes: Its sole job is to call the library to build the custom menus.
   @returns {void}
*/
function onOpen() {
  FlexLib.fCreateCodexMenu();
  FlexLib.fCreateDesignerMenu();
} // End function onOpen

/* function fMenuTagVerification
   Purpose: The local trigger function called by the "Tag Verification" menu item.
   Assumptions: None.
   Notes: This function acts as a simple pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuTagVerification() {
  FlexLib.run('TagVerification');
} // End function fMenuTagVerification

/* function fMenuToggleVisibility
   Purpose: Local trigger for the "Show/Hide All" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuToggleVisibility() {
  FlexLib.run('ToggleVisibility');
} // End function fMenuToggleVisibility


/* function fMenuClearProperties
   Purpose: Local trigger for the "Clear PropertiesService" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuClearProperties() {
  FlexLib.run('ClearProperties');
} // End function fMenuClearProperties

/* function fMenuTest
   Purpose: Local trigger for the "Test" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuTest() {
  FlexLib.run('Test');
} // End function fMenuTest


/* function fMenuCreateLatestCharacter
   Purpose: Local trigger for the "Create New Character > Latest Version" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuCreateLatestCharacter() {
  FlexLib.run('CreateLatestCharacter');
} // End function fMenuCreateLatestCharacter

/* function fMenuCreateLegacyCharacter
   Purpose: Local trigger for the "Create New Character > Older Legacy Version" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuCreateLegacyCharacter() {
  FlexLib.run('CreateLegacyCharacter');
} // End function fMenuCreateLegacyCharacter

/* function fMenuRenameCharacter
   Purpose: Local trigger for the "Rename Character" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuRenameCharacter() {
  FlexLib.run('RenameCharacter');
} // End function fMenuRenameCharacter

/* function fMenuDeleteCharacter
   Purpose: Local trigger for the "Delete Character(s)" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuDeleteCharacter() {
  FlexLib.run('DeleteCharacter');
} // End function fMenuDeleteCharacter

/* function fMenuCreateCharacter
   Purpose: Local trigger for the "Create New Character" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuCreateCharacter() {
  FlexLib.run('CreateCharacter');
} // End function fMenuCreateCharacter
