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
   @param {string} [sheetToActivate] - The optional name of the sheet to activate before running the command.
   @returns {void}
*/
function run(command, sheetToActivate) {
  try {
    if (sheetToActivate) {
      fActivateSheetByName(sheetToActivate);
    }

    const commandMap = {
      // --- THIS IS THE FIX ---
      InvalidateGameCache: () => fInvalidateSheetCache('CS', 'Game'),
      ShareCustomLists: fShareCustomLists,
      RenameCustomList: fRenameCustomList,
      DeleteCustomList: fDeleteCustomList,
      DeleteSelectedPowers: fDeleteSelectedPowers,
      VerifyAndPublish: fVerifyAndPublish,
      VerifyAndPublishMagicItems: fVerifyAndPublishMagicItems,
      DeleteSelectedMagicItems: fDeleteSelectedMagicItems,
      CreateCustomList: fCreateNewCustomList,
      SyncPowerChoices: fUpdatePowerTablesList,
      SyncMagicItemChoices: fUpdateMagicItemChoices,
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
      BuildMagicItems: fBuildMagicItems,
      FilterPowers: fFilterPowers,
      FilterMagicItems: fFilterMagicItems,
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