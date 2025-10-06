/* global FlexLib, SpreadsheetApp, PropertiesService */

const SCRIPT_INITIALIZED_KEY = 'SCRIPT_INITIALIZED';

/* function onOpen
   Purpose: Simple trigger that runs automatically when the spreadsheet is opened.
   Assumptions: None.
   Notes: Builds the full menu if authorized, otherwise provides an activation option.
   @returns {void}
*/
function onOpen() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const isInitialized = scriptProperties.getProperty(SCRIPT_INITIALIZED_KEY);

  if (isInitialized) {
    FlexLib.fCreateDesignerMenu('Tables');
  } else {
    SpreadsheetApp.getUi()
      .createMenu('*** Flex ***')
      .addItem('▶️ Activate Flex Menus', 'fActivateMenus')
      .addToUi();
  }
} // End function onOpen

/* function fActivateMenus
   Purpose: Runs the first-time authorization and menu setup.
   Assumptions: Triggered by a user clicking the 'Activate' menu item.
   Notes: This function's execution by a user triggers the Google Auth prompt if needed.
   @returns {void}
*/
function fActivateMenus() {
  // This line is the key. Calling any service that needs permissions inside a function
  // triggered by a menu click will initiate the authorization flow if the user
  // has not yet approved the script for this document.
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty(SCRIPT_INITIALIZED_KEY, 'true');

  // Let the user know it worked and what to do next.
  SpreadsheetApp.getUi().alert(
    '✅ Success!',
    'The script has been authorized. Please refresh the page. The full menus will appear automatically from now on.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
} // End function fActivateMenus

/* function fMenuPlaceholder
   Purpose: Local trigger for placeholder menu items.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuPlaceholder() {
  FlexLib.run('ShowPlaceholder');
} // End function fMenuPlaceholder

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