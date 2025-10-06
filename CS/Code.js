/* global FlexLib, SpreadsheetApp */

// --- Session Caches for High-Speed Performance ---
let powerDataCache = null; // Caches the filtered power data.
let csHeaderCache = null; // Caches the Character Sheet header row.

/* function onOpen
   Purpose: Simple trigger that runs automatically when the spreadsheet is opened.
   Assumptions: None.
   Notes: Its sole job is to call the library to build the custom menu.
   @returns {void}
*/
function onOpen() {
  FlexLib.fCreateGenericMenus('CS');
} // End function onOpen

/* function onEdit
   Purpose: A simple trigger that auto-populates power details from a high-speed session cache when a power is selected from a dropdown.
   Assumptions: The <PowerDataCache> sheet exists. The <Game> sheet is tagged correctly.
   Notes: This is the final, optimized auto-formatter. The first run in a session is slow; subsequent runs are instant.
   @param {GoogleAppsScript.Events.SheetsOnEdit} e - The event object passed by the trigger.
   @returns {void}
*/
function onEdit(e) {
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const col = e.range.getColumn();
  const row = e.range.getRow();
  const selectedValue = e.value;

  if (sheetName !== 'Game') {
    return;
  }

  // 1. Use the header cache. If it's empty, populate it once.
  if (!csHeaderCache) {
    csHeaderCache = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }
  const csHeader = csHeaderCache;

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

  // 2. Use the power data cache. If it's empty, populate it once.
  if (!powerDataCache) {
    const cacheSheet = e.source.getSheetByName('PowerDataCache');
    if (!cacheSheet) return;

    const allCachedPowers = cacheSheet.getDataRange().getValues();
    const cacheHeader = allCachedPowers.shift(); // Remove header row

    const cacheColMap = {
      dropdown: cacheHeader.indexOf('DropDown'),
      usage: cacheHeader.indexOf('Usage'),
      action: cacheHeader.indexOf('Action'),
      name: cacheHeader.indexOf('Power'),
      effect: cacheHeader.indexOf('Effect'),
    };

    powerDataCache = new Map();
    allCachedPowers.forEach(pRow => {
      const key = pRow[cacheColMap.dropdown];
      const value = {
        usage: pRow[cacheColMap.usage],
        action: pRow[cacheColMap.action],
        name: pRow[cacheColMap.name],
        effect: pRow[cacheColMap.effect],
      };
      powerDataCache.set(key, value);
    });
  }

  const powerData = powerDataCache.get(selectedValue);

  if (!powerData) return;

  // 3. Write the details to the correct adjacent cells
  sheet.getRange(row, csHeader.indexOf(csTargetTags.usage) + 1).setValue(powerData.usage);
  sheet.getRange(row, csHeader.indexOf(csTargetTags.action) + 1).setValue(powerData.action);
  sheet.getRange(row, csHeader.indexOf(csTargetTags.name) + 1).setValue(powerData.name);
  sheet.getRange(row, csHeader.indexOf(csTargetTags.effect) + 1).setValue(powerData.effect);
} // End function onEdit


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