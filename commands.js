/* Tarifast v9 minimal event runtime test.
 * Match the Office initialization used by Microsoft's external-recipient sample.
 * Keep event association outside initialization for classic Outlook on Windows.
 */
function tarifastDiagnostic(message) {
  try {
    console.log("Tarifast diagnostic v2: " + message);
  } catch (_) {
    // Diagnostics must never interrupt event handling.
  }
}

tarifastDiagnostic("1 - script executing");

Office.initialize = function () {
  tarifastDiagnostic("3 - Office initialization callback reached");
};

function tarifastV9RecipientsChanged(event) {
  tarifastDiagnostic("4 - recipient event fired");

  try {
    event.completed();
  } catch (_) {
    tarifastDiagnostic("event.completed threw an error");
  }
}

Office.actions.associate(
  "tarifastV9RecipientsChanged",
  tarifastV9RecipientsChanged
);
tarifastDiagnostic("2 - handler association returned");
