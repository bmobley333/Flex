/* global FlexLib */

/* function onOpen
   Purpose: Simple trigger that runs automatically when the spreadsheet is opened.
   Assumptions: None.
   Notes: Its sole job is to call the library to build the custom menu.
   @returns {void}
*/
function onOpen() {
  FlexLib.fCreateStandardMenus();
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

/* function fMenuTest
   Purpose: Local trigger for the "Test" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuTest() {
  FlexLib.run('Test');
} // End function fMenuTest

/* function fMenuUpdatePowerTables
   Purpose: Local trigger for the "Update Power Tables" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuUpdatePowerTables() {
  FlexLib.run('UpdatePowerTables');
} // End function fMenuUpdatePowerTables

/* function fMenuFilterPowers
   Purpose: Local trigger for the "Filter Powers" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuFilterPowers() {
  FlexLib.run('FilterPowers');
} // End function fMenuFilterPowers

/* function onEdit
   Purpose: A simple trigger that auto-populates power details from a local cache when a power is selected from a dropdown.
   Assumptions: The <PowerDataCache> sheet exists and is populated. The <Game> sheet is tagged correctly.
   Notes: This is the auto-formatter that provides a seamless UX for the player.
   @param {GoogleAppsScript.Events.SheetsOnEdit} e - The event object passed by the trigger.
   @returns {void}
*/
function onEdit(e) {
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const col = e.range.getColumn();
  const row = e.range.getRow();
  const selectedValue = e.value;

  // --- Guard Clause ---
  if (sheetName !== 'Game') {
    return;
  }

  const csHeader = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const editedColTag = csHeader[col - 1];
  if (!editedColTag || !editedColTag.startsWith('PowerDropDown')) {
    return;
  }

  const tagIndex = editedColTag.replace('PowerDropDown', '');
  const csTargetTags = {
    usage: `PowerUsage${tagIndex}`,
    action: `PowerAction${tagIndex}`,
    name: `PowerName${tagIndex}`,
    effect: `PowerEffect${tagIndex}`,
  };

  if (!selectedValue) {
    Object.values(csTargetTags).forEach(tag => {
      const targetColIndex = csHeader.indexOf(tag);
      if (targetColIndex !== -1) {
        sheet.getRange(row, targetColIndex + 1).clearContent();
      }
    });
    return;
  }

  // 1. Read data from the LOCAL <PowerDataCache> sheet
  const cacheSheet = e.source.getSheetByName('PowerDataCache');
  if (!cacheSheet) return; // Fail silently if cache is missing

  const allCachedPowers = cacheSheet.getDataRange().getValues();
  const cacheHeader = allCachedPowers[0];

  const cacheColMap = {
    dropdown: cacheHeader.indexOf('DropDown'),
    usage: cacheHeader.indexOf('Usage'),
    action: cacheHeader.indexOf('Action'),
    name: cacheHeader.indexOf('Power'), // <-- Corrected from 'AbilityName'
    effect: cacheHeader.indexOf('Effect'),
  };

  // 2. Find the selected power in the cached data
  const powerData = allCachedPowers.find(pRow => pRow[cacheColMap.dropdown] === selectedValue);

  if (!powerData) return; // Power not found in cache

  // 3. Write the details to the correct adjacent cells on the CS
  sheet.getRange(row, csHeader.indexOf(csTargetTags.usage) + 1).setValue(powerData[cacheColMap.usage]);
  sheet.getRange(row, csHeader.indexOf(csTargetTags.action) + 1).setValue(powerData[cacheColMap.action]);
  sheet.getRange(row, csHeader.indexOf(csTargetTags.name) + 1).setValue(powerData[cacheColMap.name]);
  sheet.getRange(row, csHeader.indexOf(csTargetTags.effect) + 1).setValue(powerData[cacheColMap.effect]);
} // End function onEdit