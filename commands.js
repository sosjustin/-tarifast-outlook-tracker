/* ============================================================
   TARIFAST EMAIL TRACKER
   VERSION 10 — EVENT ACTIVATION DIAGNOSTIC

   PURPOSE
   ------------------------------------------------------------
   Determine whether Outlook dispatches:

   1. OnNewMessageCompose
   2. OnMessageRecipientsChanged

   NO:
   - Wix calls
   - fetch()
   - tracking logic
   - message modification
   - OnMessageSend
============================================================ */

function tarifastDiagnostic(message) {
  try {
    console.log(
      "TARIFAST V10 DIAGNOSTIC | " +
      new Date().toISOString() +
      " | " +
      message
    );
  } catch (_) {
    // Diagnostics must never interrupt event handling.
  }
}


/* ============================================================
   SCRIPT LOAD DIAGNOSTIC
============================================================ */

tarifastDiagnostic("SCRIPT EXECUTING");


/* ============================================================
   OFFICE INITIALIZATION DIAGNOSTIC

   This is informational only.
   Event-based activation does not depend on Office.initialize.
============================================================ */

Office.initialize = function () {
  tarifastDiagnostic("OFFICE INITIALIZE CALLBACK");
};


/* ============================================================
   NEW MESSAGE COMPOSE
============================================================ */

function tarifastV10Compose(event) {

  tarifastDiagnostic("NEW MESSAGE COMPOSE FIRED");

  try {
    event.completed();
  } catch (error) {
    tarifastDiagnostic(
      "NEW MESSAGE COMPOSE event.completed ERROR: " +
      String(error)
    );
  }
}


/* ============================================================
   RECIPIENTS CHANGED
============================================================ */

function tarifastV10RecipientsChanged(event) {

  tarifastDiagnostic("RECIPIENT CHANGE FIRED");

  try {
    event.completed();
  } catch (error) {
    tarifastDiagnostic(
      "RECIPIENT CHANGE event.completed ERROR: " +
      String(error)
    );
  }
}


/* ============================================================
   EVENT ASSOCIATIONS
============================================================ */

try {

  Office.actions.associate(
    "tarifastV10Compose",
    tarifastV10Compose
  );

  tarifastDiagnostic(
    "COMPOSE HANDLER ASSOCIATION RETURNED"
  );

} catch (error) {

  tarifastDiagnostic(
    "COMPOSE HANDLER ASSOCIATION ERROR: " +
    String(error)
  );

}


try {

  Office.actions.associate(
    "tarifastV10RecipientsChanged",
    tarifastV10RecipientsChanged
  );

  tarifastDiagnostic(
    "RECIPIENT HANDLER ASSOCIATION RETURNED"
  );

} catch (error) {

  tarifastDiagnostic(
    "RECIPIENT HANDLER ASSOCIATION ERROR: " +
    String(error)
  );

}
