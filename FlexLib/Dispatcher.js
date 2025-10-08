/* global fShowMessage, fVerifyActiveSheetTags */
/* exported run */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Command Dispatcher
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function run
   Purpose: Acts as the central dispatcher for all commands initiated from a local sheet script.
   Assumptions: The command string passed matches a key in the commandMap.
   Notes: This provides a single entry point and a master try/catch for robust error handling.
   @param {string} command - The unique identifier for the command to execute.
   @returns {void}
*/
function run(command) {
  try {
    const commandMap = {
      CreateCustomList: fCreateNewCustomList, // <-- ADDED
      ShareMyAbilities: fShareMyAbilities,
      SyncPowerChoices: fUpdatePowerTablesList,
      AddNewCustomSource: fAddNewCustomSource,
      InitialSetup: fInitialSetup,
      TagVerification: fVerifyActiveSheetTags,
      ToggleVisibility: fToggleDesignerVisibility,
      TrimSheet: fTrimSheet,
      Test: fTestIdManagement,
      CreateLatestCharacter: fCreateLatestCharacter,
      CreateLegacyCharacter: fCreateLegacyCharacter,
      RenameCharacter: fRenameCharacter,
      DeleteCharacter: fDeleteCharacter,
      ShowPlaceholder: fShowPlaceholderMessage,
      BuildPowers: fBuildPowers,
      FilterPowers: fFilterPowers,
      PrepGameForPaper: fPrepGameForPaper,
    };

    if (commandMap[command]) {
      commandMap[command]();
    } else {
      throw new Error(`Unknown command received: ${command}`);
    }
  } catch (e) {
    console.error(e);
    fShowMessage('❌ Error', e.message);
  }
} // End function run