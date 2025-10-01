/* global g */
/* exported g */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Global Constants & State
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

const g = {
  // Developer Info
  ADMIN_EMAIL: 'metascapegame@gmail.com',
  CURRENT_VERSION: '3',
  // The master source of truth for all game versions
  MASTER_VER_ID: '1zlSL-B0k3vPen5EEG0-DPZXO9dwes_hB4CdcL9ngolU',
  MASTER_VER_INFO: {
    version: 'N/A',
    ssabbr: 'Ver',
    ssid: '1zlSL-B0k3vPen5EEG0-DPZXO9dwes_hB4CdcL9ngolU',
    ssfullname: 'Versions',
  },

  // Object to cache the mapping of Version -> Abbr -> full data object
  sheetIDs: {},

  // Object structures for sheet data (arrays and tag maps)
  Ver: {},
  Codex: {},
  CS: {},
  DB: {},
}; // End const g