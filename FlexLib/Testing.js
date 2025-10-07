/* global g, fBuildTagMaps, fShowMessage */
/* exported fTestTagMaps */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Testing Utilities
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


/* function fClearAllScriptProperties
   Purpose: A utility to completely clear the script's property cache.
   Assumptions: To be run manually from the script editor for testing purposes.
   Notes: This is essential for re-testing first-run experiences.
   @returns {void}
*/
function fClearAllScriptProperties() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  fShowMessage('✅ Success', 'All script properties have been cleared.');
} // End function fClearAllScriptProperties


/* function fTestIdManagement
   Purpose: A test function to verify that the ID caching system is working.
   Assumptions: The 'Codex' spreadsheet has a 'Versions' sheet with data for a 'DB' entry.
   Notes: Displays all cached info for the 'DB' sheet abbreviation.
   @returns {void}
*/
function fTestIdManagement() {
  fGetSheetId(g.CURRENT_VERSION, 'DB'); // This triggers the caching logic

  const dbInfo = g.sheetIDs[g.CURRENT_VERSION]['DB'];

  if (!dbInfo) {
    throw new Error("Could not retrieve info for ssAbbr 'DB'. Check the 'Versions' sheet.");
  }

  let message = '✅ ID Cache Loaded Successfully!\n\n';
  message += `Version: ${dbInfo.version}\n`;
  message += `Full Name: ${dbInfo.ssfullname}\n`;
  message += `Abbreviation: ${dbInfo.ssabbr}\n`;
  message += `ID: ${dbInfo.ssid}`;

  fShowMessage('Test Results', message);
} // End function fTestIdManagement