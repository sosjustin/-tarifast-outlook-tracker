/* ============================================================
   TARIFAST EMAIL TRACKER
   VERSION 9 — MINIMAL EVENT RUNTIME TEST

   PURPOSE:
   Prove that Outlook can invoke the
   OnMessageRecipientsChanged handler successfully.

   IMPORTANT:
   - No Wix calls
   - No fetch()
   - No recipient lookup
   - No tracking logic
   - No promises
   - No message modification

   The ONLY operation performed when the event fires is
   event.completed().
============================================================ */


/* ============================================================
   RECIPIENT CHANGE HANDLER

   This function name MUST exactly match the FunctionName
   declared in the v9 Outlook manifest:

   tarifastV9RecipientsChanged
============================================================ */

function tarifastV9RecipientsChanged(event) {

  try {

    event.completed();

  } catch (_) {

    /*
     * Nothing else should execute.
     * This is intentionally a minimal runtime test.
     */

  }

}


/* ============================================================
   MICROSOFT EVENT ASSOCIATION

   Associates the manifest FunctionName with the JavaScript
   handler above.
============================================================ */

Office.actions.associate(
  "tarifastV9RecipientsChanged",
  tarifastV9RecipientsChanged
);
