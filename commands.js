/* ============================================================
   TARIFAST EMAIL TRACKER
   VERSION 9 CONTROL TEST

   PURPOSE:
   Prove Microsoft's OnMessageRecipientsChanged event
   actually invokes the registered JavaScript handler.

   NO PRODUCTION TRACKING LOGIC IS PRESENT IN THIS TEST.
============================================================ */


const TRACKING_CREATE_URL =
  "https://www.tarifastops.com/_functions/emailTrackingCreate";


/* ============================================================
   RECIPIENT CHANGE HANDLER

   IMPORTANT:
   The function name is intentionally NEW for v9.

   Manifest:
     tarifastV9RecipientsChanged

   JavaScript association:
     tarifastV9RecipientsChanged

   This eliminates any ambiguity involving the old cached
   handler name used by v8.
============================================================ */

function tarifastV9RecipientsChanged(event) {

  try {

    fetch(
      TRACKING_CREATE_URL,
      {
        method: "POST",

        headers: {
          "Content-Type": "application/json"
        },

        body: JSON.stringify({

          recipientEmail:
            "jlesperance@tarifastops.com",

          recipientName:
            "TARIFAST V9 EVENT CONTROL",

          company:
            "Tarifast",

          subject:
            "V9_HANDLER_ACTUALLY_FIRED",

          messageId:
            "v9-handler-" +
            Date.now()

        })

      }
    )

    .catch(function () {

      /*
       * Never allow diagnostic failure to affect Outlook.
       */

    })

    .then(function () {

      try {
        event.completed();
      } catch (_) {}

    });


  } catch (_) {

    /*
     * Absolute fail-open.
     */

    try {
      event.completed();
    } catch (_) {}

  }

}


/* ============================================================
   MICROSOFT EVENT ASSOCIATION

   This name MUST exactly match FunctionName in manifest v9.
============================================================ */

Office.actions.associate(
  "tarifastV9RecipientsChanged",
  tarifastV9RecipientsChanged
);
