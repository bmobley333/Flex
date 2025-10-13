/* global FlexLib, SpreadsheetApp */

// --- Session Caches for High-Speed Performance ---
let powerDataCache = null; // Caches the filtered power data.
let magicItemDataCache = null; // Caches the filtered magic item data.
let csHeaderCache = null; // Caches the Character Sheet header row.

const SCRIPT_INITIALIZED_KEY = 'SCRIPT_INITIALIZED';


/* function onOpen
   Purpose: Simple trigger that runs automatically when the spreadsheet is opened.
   Assumptions: None.
   Notes: Builds menus based on authorization status and user identity (player vs. designer).
   @returns {void}
*/
function onOpen() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const isInitialized = scriptProperties.getProperty(SCRIPT_INITIALIZED_KEY);

  if (isInitialized) {
    // Always create the main player menu.
    FlexLib.fCreateFlexMenu();

    // Get the globals object from the library.
    const g = FlexLib.getGlobals();

    // Only show the Designer menu if the user is the admin.
    if (Session.getActiveUser().getEmail() === g.ADMIN_EMAIL) {
      FlexLib.fCreateDesignerMenu('CS');
    }
  } else {
    SpreadsheetApp.getUi()
      .createMenu('💪 Flex')
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


/* function onEdit
   Purpose: A simple trigger that auto-populates details from a high-speed session cache when an item is selected from a dropdown.
   Assumptions: The appropriate DataCache sheet exists. The <Game> sheet is tagged correctly.
   Notes: This is the optimized auto-formatter. First run in a session is slow; subsequent runs are instant.
   @param {GoogleAppsScript.Events.SheetsOnEdit} e - The event object passed by the trigger.
   @returns {void}
*/
function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== 'Game') return;

  try {
    const col = e.range.getColumn();
    const row = e.range.getRow();
    const selectedValue = e.value;

    if (!csHeaderCache) {
      csHeaderCache = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    }
    const csHeader = csHeaderCache;
    const editedColTag = csHeader[col - 1];

    if (!editedColTag || !editedColTag.includes('DropDown')) {
      return;
    }

    let tagPrefix = null;
    let data = null;

    // 1. Check existing caches first
    if (powerDataCache && powerDataCache.has(selectedValue)) {
      tagPrefix = 'Power';
      data = powerDataCache.get(selectedValue);
    } else if (magicItemDataCache && magicItemDataCache.has(selectedValue)) {
      tagPrefix = 'MagicItem';
      data = magicItemDataCache.get(selectedValue);
    }

    // 2. If not in cache, build and check again
    if (!data && selectedValue) {
      if (!powerDataCache) powerDataCache = FlexLib.fBuildCache('PowerDataCache', 'Power');
      if (powerDataCache.has(selectedValue)) {
        tagPrefix = 'Power';
        data = powerDataCache.get(selectedValue);
      } else {
        if (!magicItemDataCache) magicItemDataCache = FlexLib.fBuildCache('MagicItemDataCache', 'Name');
        if (magicItemDataCache.has(selectedValue)) {
          tagPrefix = 'MagicItem';
          data = magicItemDataCache.get(selectedValue);
        }
      }
    }

    // 3. Find the number from the edited column tag
    let tagNumber = '1';
    if (editedColTag) {
      const match = editedColTag.match(/\d+$/);
      if (match) tagNumber = match[0];
    }

    const findColumnIndexByTag = (partialTag) => {
      return csHeader.findIndex(headerTag => headerTag.includes(partialTag));
    };

    // 4. Clear or populate cells
    if (!selectedValue || !data) {
      const allPossibleTags = [`PowerUsage${tagNumber}`, `PowerAction${tagNumber}`, `PowerName${tagNumber}`, `PowerEffect${tagNumber}`, `MagicItemUsage${tagNumber}`, `MagicItemAction${tagNumber}`, `MagicItemName${tagNumber}`, `MagicItemEffect${tagNumber}`];
      allPossibleTags.forEach(tag => {
        const targetColIndex = findColumnIndexByTag(tag);
        if (targetColIndex !== -1) {
          sheet.getRange(row, targetColIndex + 1).clearContent();
        }
      });
      return;
    }

    const targetTags = {
      usage: `${tagPrefix}Usage${tagNumber}`,
      action: `${tagPrefix}Action${tagNumber}`,
      name: `${tagPrefix}Name${tagNumber}`,
      effect: `${tagPrefix}Effect${tagNumber}`,
    };

    const usageCol = findColumnIndexByTag(targetTags.usage);
    const actionCol = findColumnIndexByTag(targetTags.action);
    const nameCol = findColumnIndexByTag(targetTags.name);
    const effectCol = findColumnIndexByTag(targetTags.effect);

    if (usageCol !== -1) sheet.getRange(row, usageCol + 1).setValue(data.usage);
    if (actionCol !== -1) sheet.getRange(row, actionCol + 1).setValue(data.action);
    if (nameCol !== -1) sheet.getRange(row, nameCol + 1).setValue(data.name);
    if (effectCol !== -1) sheet.getRange(row, effectCol + 1).setValue(data.effect);

  } catch (e) {
    console.error(`❌ CRITICAL ERROR in onEdit: ${e.message}\n${e.stack}`);
  }
} // End function onEdit

/* function buttonFilterPowers
   Purpose: Local trigger for a button, mimics the "Filter Powers" menu item.
   Assumptions: None.
   Notes: Assign this function name to a button in the sheet to trigger the FilterPowers command.
   @returns {void}
*/
function buttonFilterPowers() {
  FlexLib.run('FilterPowers');
} // End function buttonFilterPowers


/* function onChange
   Purpose: An installable trigger that invalidates session caches when the sheet's structure changes.
   Assumptions: This trigger is manually installed for the spreadsheet.
   Notes: This protects against data corruption if a user inserts/deletes rows or columns.
   @param {GoogleAppsScript.Events.SheetsOnChange} e - The event object passed by the trigger.
   @returns {void}
*/
function onChange(e) {
  const structuralChanges = ['INSERT_ROW', 'REMOVE_ROW', 'INSERT_COLUMN', 'REMOVE_COLUMN'];
  if (structuralChanges.includes(e.changeType)) {
    // A structural change was made, so we must invalidate our caches.
    powerDataCache = null;
    csHeaderCache = null;
    console.log('Cache invalidated due to structural sheet change.');
  }
} // End function onChange


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

/* function fMenuSyncPowerChoices
   Purpose: Local trigger for the "Sync Power Choices" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuSyncPowerChoices() {
  FlexLib.run('SyncPowerChoices');
} // End function fMenuSyncPowerChoices

/* function fMenuFilterPowers
   Purpose: Local trigger for the "Filter Powers" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuFilterPowers() {
  FlexLib.run('FilterPowers');
} // End function fMenuFilterPowers

/* function fMenuSyncMagicItemChoices
   Purpose: Local trigger for the "Sync Magic Item Choices" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuSyncMagicItemChoices() {
  FlexLib.run('SyncMagicItemChoices');
} // End function fMenuSyncMagicItemChoices

/* function fMenuFilterMagicItems
   Purpose: Local trigger for the "Filter Magic Items" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuFilterMagicItems() {
  FlexLib.run('FilterMagicItems');
} // End function fMenuFilterMagicItems

/* function fMenuPrepGameForPaper
   Purpose: Local trigger for the "Copy CS <Game> to <Paper>" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuPrepGameForPaper() {
  FlexLib.run('PrepGameForPaper');
} // End function fMenuPrepGameForPaper