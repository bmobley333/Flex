/* global FlexLib, PropertiesService, SpreadsheetApp, Session */

const SCRIPT_INITIALIZED_KEY = 'CODEX_INITIALIZED';

/* function onOpen
   Purpose: Simple trigger that runs automatically when the spreadsheet is opened.
   Assumptions: None.
   Notes: Builds menus based on authorization status and user identity (player vs. designer).
   @returns {void}
*/
function onOpen() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const isInitialized = scriptProperties.getProperty(SCRIPT_INITIALIZED_KEY);
  const g = FlexLib.getGlobals();
  const isAdmin = Session.getActiveUser().getEmail() === g.ADMIN_EMAIL;

  if (isInitialized || isAdmin) {
    // If the script is initialized OR the user is the admin, create the full menus.
    FlexLib.fCreateCodexMenu();
    if (isAdmin) {
      FlexLib.fCreateDesignerMenu('Codex');
    }
  } else {
    // If not initialized, show a simple menu to activate the script.
    SpreadsheetApp.getUi()
      .createMenu('*** Flex ***')
      .addItem('▶️ Activate Flex Menus', 'fActivateCodex')
      .addToUi();
  }
} // End function onOpen

/* function fActivateCodex
   Purpose: Runs the first-time authorization and setup, then prompts the user to refresh.
   Assumptions: Triggered by a user clicking the 'Activate' menu item.
   Notes: This function's execution triggers the Google Auth prompt and the one-time setup.
   @returns {void}
*/
function fActivateCodex() {
  // First, call the library to run the initial setup.
  FlexLib.run('InitialSetup');

  // Once setup is complete, set the property so the full menus appear next time.
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty(SCRIPT_INITIALIZED_KEY, 'true');

  // Display a consistent success message.
  const title = 'IMPORTANT - Please Refresh Browser Tab';
  const message = '✅ Success! The script has been authorized and setup is complete.\n\nPlease refresh this browser tab now to load the full custom menus.';
  SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
} // End function fActivateCodex

/* function fMenuTrimSheet
   Purpose: Local trigger for the "Trim Empty Rows/Cols" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuTrimSheet() {
  FlexLib.run('TrimSheet');
} // End function fMenuTrimSheet

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
  // This first call to the library will trigger the auth prompt if needed,
  // and then immediately display a toast once authorization is complete.
  FlexLib.fShowToast('⏳ Initializing setup...', 'Welcome to Flex');
  FlexLib.run('CreateLatestCharacter');
} // End function fMenuCreateLatestCharacter

/* function fMenuCreateLegacyCharacter
   Purpose: Local trigger for the "Create New Character > Older Legacy Version" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuCreateLegacyCharacter() {
  // This first call to the library will trigger the auth prompt if needed,
  // and then immediately display a toast once authorization is complete.
  FlexLib.fShowToast('⏳ Initializing setup...', 'Welcome to Flex');
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


/* function fMenuAddNewCustomSource
   Purpose: Local trigger for the "Add New Source..." menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuAddNewCustomSource() {
  FlexLib.run('AddNewCustomSource');
} // End function fMenuAddNewCustomSource

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
