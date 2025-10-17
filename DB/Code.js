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
    FlexLib.fCreateDesignerMenu('DB');
  } else {
    SpreadsheetApp.getUi()
      .createMenu('💪 MS3')
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
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty(SCRIPT_INITIALIZED_KEY, 'true');

  const title = 'IMPORTANT - Please Refresh Browser Tab';
  const message = '✅ Success! The script has been authorized.\n\nPlease refresh this browser tab now to load the full custom menus.';
  SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
} // End function fActivateMenus

/* function fMenuBuildSkillSets
   Purpose: Local trigger for the "Build Skill Sets" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuBuildSkillSets() {
  FlexLib.run('BuildSkillSets', 'SkillSets');
} // End function fMenuBuildSkillSets

/* function fMenuBuildMagicItems
   Purpose: Local trigger for the "Build Magic Items" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuBuildMagicItems() {
  FlexLib.run('BuildMagicItems', 'Magic Items');
} // End function fMenuBuildMagicItems

/* function fMenuBuildPowers
   Purpose: Local trigger for the "Build Powers" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuBuildPowers() {
  FlexLib.run('BuildPowers', 'Powers');
} // End function fMenuBuildPowers

/* function fMenuPlaceholder
   Purpose: Local trigger for placeholder menu items.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuPlaceholder() {
  FlexLib.run('ShowPlaceholder');
} // End function fMenuPlaceholder


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