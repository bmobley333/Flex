/* global fShowMessage */
/* exported run */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Dispatcher & Ver Logic
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function run
   Purpose: Acts as the central dispatcher for all commands initiated from a local sheet script.
   Assumptions: The command string passed matches a known command.
   Notes: This provides a single entry point and a master try/catch for robust error handling.
   @param {string} command - The unique identifier for the command to execute.
   @returns {void}
*/
function run(command) {
  try {
    // We use a simple switch statement which is very reliable in GAS.
    switch (command) {
      case 'Ver_TagVerification':
        fShowMessage('Tag Verification', '✅ Tag Verification Worked');
        break;

      default:
        throw new Error(`Unknown command received: ${command}`);
    }
  } catch (e) {
    console.error(e);
    // This ensures the error message itself is always displayed correctly.
    fShowMessage('❌ Error', e.message);
  }
} // End function run