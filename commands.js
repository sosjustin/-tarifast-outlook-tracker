/* ============================================================
   TARIFAST EMAIL TRACKER
   MINIMAL EVENT ACTIVATION TEST

   PURPOSE:
   Prove whether Outlook actually invokes
   OnMessageRecipientsChanged.

   IMPORTANT:
   - No OnMessageSend
   - No appendOnSendAsync
   - No sessionData
   - No startup diagnostics
   - No production tracking logic
   - Nothing here can block Send
============================================================ */


/* ============================================================
   RECIPIENT CHANGE EVENT
============================================================ */

function onMessageRecipientsChangedHandler(event) {

  try {

    fetch(
      "https://www.tarifastops.com/_functions/emailTrackingCreate",
      {
        method: "POST",

        headers: {
          "Content-Type": "application/json"
        },

        body: JSON.stringify({

          recipientEmail:
            "jlesperance@tarifastops.com",

          recipientName:
            "EVENT HANDLER TEST",

          company:
            "Tarifast",

          subject:
            "HANDLER_ACTUALLY_FIRED",

          messageId:
            "handler-test-" +
            Date.now()

        })

      }
    )

    .catch(function () {

      /*
       * Diagnostic request failure is ignored.
       * It must never interfere with Outlook.
       */

    })

    .finally(function () {

      /*
       * Always tell Outlook the event is complete.
       */

      try {

        event.completed();

      } catch (_) {}

    });


  } catch (_) {

    /*
     * Absolute fail-open protection.
     *
     * Even an unexpected JavaScript error must not
     * interfere with Outlook.
     */

    try {

      event.completed();

    } catch (_) {}

  }

}


/* ============================================================
   REGISTER EVENT HANDLER

   This name MUST exactly match the FunctionName declared
   by the v8 manifest:

   onMessageRecipientsChangedHandler
============================================================ */

Office.actions.associate(
  "onMessageRecipientsChangedHandler",
  onMessageRecipientsChangedHandler
);
