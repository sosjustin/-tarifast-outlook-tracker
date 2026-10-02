/* Tarifast v9 minimal event runtime test.
 * Match the Office initialization used by Microsoft's external-recipient sample.
 * Keep event association outside initialization for classic Outlook on Windows.
 */
Office.initialize = function () {};

function tarifastV9RecipientsChanged(event) {
  try {
    event.completed();
  } catch (_) {
    // Preserve the existing minimal test's behavior.
  }
}

Office.actions.associate(
  "tarifastV9RecipientsChanged",
  tarifastV9RecipientsChanged
);
