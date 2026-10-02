const TRACKING_CREATE_URL =
  "https://www.tarifastops.com/_functions/emailTrackingCreate";

const TRACKING_SESSION_KEY =
  "tarifastEmailTrackingPixelUrl";


/* =========================================================
   DIAGNOSTIC HELPER

   TEMPORARY:
   Writes diagnostic stages to the Wix EmailTracking
   collection so we can determine exactly where Outlook
   event activation succeeds or fails.

   All diagnostics are fail-open.
========================================================= */

async function writeDiagnostic(stage) {
  try {

    await fetch(
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
            "OUTLOOK JS DIAGNOSTIC",

          company:
            "Tarifast",

          subject:
            String(stage || "UNKNOWN_DIAGNOSTIC"),

          messageId:
            "outlook-js-" +
            Date.now() +
            "-" +
            Math.random()
              .toString(36)
              .slice(2, 8)
        })
      }
    );

  } catch (_) {

    /*
     * Diagnostics must never interfere with Outlook.
     */

  }
}


/* =========================================================
   DIAGNOSTIC STAGE 1

   If this record appears, commands.js itself loaded and
   began executing.
========================================================= */

writeDiagnostic(
  "OUTLOOK_COMMANDS_JS_LOADED"
);


/* =========================================================
   OFFICE.JS PROMISE WRAPPERS
========================================================= */

function officeGetAsync(target) {

  return new Promise((resolve, reject) => {

    target.getAsync((result) => {

      if (
        result.status ===
        Office.AsyncResultStatus.Succeeded
      ) {

        resolve(
          result.value
        );

      } else {

        reject(
          result.error ||
          new Error(
            "Office getAsync failed"
          )
        );

      }

    });

  });

}


function bodyGetTypeAsync(body) {

  return new Promise((resolve, reject) => {

    body.getTypeAsync((result) => {

      if (
        result.status ===
        Office.AsyncResultStatus.Succeeded
      ) {

        resolve(
          result.value
        );

      } else {

        reject(
          result.error ||
          new Error(
            "body.getTypeAsync failed"
          )
        );

      }

    });

  });

}


function sessionGetAsync(
  sessionData,
  key
) {

  return new Promise((resolve) => {

    sessionData.getAsync(
      key,
      (result) => {

        if (
          result.status ===
          Office.AsyncResultStatus.Succeeded
        ) {

          resolve(
            result.value || ""
          );

        } else {

          /*
           * Session data failure should never stop tracking
           * from attempting to continue.
           */

          resolve("");

        }

      }
    );

  });

}


function sessionSetAsync(
  sessionData,
  key,
  value
) {

  return new Promise((resolve) => {

    sessionData.setAsync(
      key,
      value,
      () => resolve()
    );

  });

}


function appendOnSendAsync(
  body,
  content,
  coercionType
) {

  return new Promise((resolve, reject) => {

    body.appendOnSendAsync(
      content,
      {
        coercionType
      },
      (result) => {

        if (
          result.status ===
          Office.AsyncResultStatus.Succeeded
        ) {

          resolve();

        } else {

          reject(
            result.error ||
            new Error(
              "appendOnSendAsync failed"
            )
          );

        }

      }
    );

  });

}


/* =========================================================
   RECIPIENT
========================================================= */

function firstRecipient(recipients) {

  if (
    !Array.isArray(recipients) ||
    recipients.length === 0
  ) {

    return null;

  }


  const recipient =
    recipients[0] || {};


  const email =
    String(
      recipient.emailAddress || ""
    ).trim();


  if (!email) {

    return null;

  }


  return {

    email,

    name:
      String(
        recipient.displayName || ""
      ).trim()

  };

}


/* =========================================================
   CLIENT-SIDE MESSAGE ID
========================================================= */

function makeClientMessageId() {

  return [

    "tarifast",

    Date.now()
      .toString(36),

    Math.random()
      .toString(36)
      .slice(2, 12)

  ].join("-");

}


/* =========================================================
   CREATE TRACKING RECORD IN WIX
========================================================= */

async function createTrackingRecord(
  recipient,
  subject
) {

  const controller =
    typeof AbortController !== "undefined"
      ? new AbortController()
      : null;


  const timeoutId =
    setTimeout(
      () => {

        if (controller) {
          controller.abort();
        }

      },
      5000
    );


  try {

    const response =
      await fetch(
        TRACKING_CREATE_URL,
        {
          method: "POST",

          headers: {
            "Content-Type":
              "application/json"
          },

          body: JSON.stringify({

            recipientEmail:
              recipient.email,

            recipientName:
              recipient.name,

            company:
              "",

            subject:
              String(
                subject || ""
              ).trim(),

            messageId:
              makeClientMessageId()

          }),

          signal:
            controller
              ? controller.signal
              : undefined
        }
      );


    if (!response.ok) {

      throw new Error(
        "Tracking endpoint returned HTTP " +
        response.status
      );

    }


    const data =
      await response.json();


    if (
      !data ||
      data.success !== true ||
      !data.pixelUrl
    ) {

      throw new Error(
        "Tracking endpoint did not return a pixel URL"
      );

    }


    return String(
      data.pixelUrl
    );

  } finally {

    clearTimeout(
      timeoutId
    );

  }

}


/* =========================================================
   RECIPIENT CHANGE EVENT

   This event prepares the tracking pixel while the message
   is being composed.

   Tracking failure must NEVER prevent an email from being
   sent.
========================================================= */

async function onMessageRecipientsChangedHandler(
  event
) {

  /*
   * DIAGNOSTIC STAGE 3
   *
   * This occurs before any mailbox/item operations.
   *
   * If this appears, Outlook successfully invoked the
   * function associated with OnMessageRecipientsChanged.
   */

  await writeDiagnostic(
    "OUTLOOK_RECIPIENT_HANDLER_ENTERED"
  );


  try {

    const item =
      Office.context.mailbox.item;


    if (
      !item ||
      !item.to ||
      !item.subject ||
      !item.body ||
      !item.sessionData ||
      typeof item.body.appendOnSendAsync !==
        "function" ||
      typeof item.body.getTypeAsync !==
        "function"
    ) {

      return;

    }


    /* =====================================================
       DUPLICATE PROTECTION

       Only prepare one tracking pixel for this compose
       session.
    ===================================================== */

    const existingPixelUrl =
      await sessionGetAsync(
        item.sessionData,
        TRACKING_SESSION_KEY
      );


    if (existingPixelUrl) {

      return;

    }


    /* =====================================================
       GET CURRENT TO RECIPIENTS
    ===================================================== */

    const recipients =
      await officeGetAsync(
        item.to
      );


    const recipient =
      firstRecipient(
        recipients
      );


    if (!recipient) {

      return;

    }


    /* =====================================================
       GET OUTLOOK BODY FORMAT
    ===================================================== */

    const bodyFormat =
      await bodyGetTypeAsync(
        item.body
      );


    /* =====================================================
       GET SUBJECT

       Subject is metadata only.

       Failure to retrieve the subject does not stop
       tracking.
    ===================================================== */

    let subject = "";


    try {

      subject =
        await officeGetAsync(
          item.subject
        );

    } catch (_) {

      subject = "";

    }


    /* =====================================================
       CREATE WIX TRACKING RECORD
    ===================================================== */

    const pixelUrl =
      await createTrackingRecord(
        recipient,
        subject
      );


    /* =====================================================
       TRACKING PIXELS REQUIRE HTML
    ===================================================== */

    if (
      bodyFormat !==
      Office.CoercionType.Html
    ) {

      return;

    }


    /* =====================================================
       ESCAPE PIXEL URL FOR HTML ATTRIBUTE
    ===================================================== */

    const safePixelUrl =
      pixelUrl
        .replace(
          /&/g,
          "&amp;"
        )
        .replace(
          /"/g,
          "&quot;"
        );


    const trackingContent =
      '<img src="' +
      safePixelUrl +
      '" width="1" height="1" alt="" />';


    /* =====================================================
       REGISTER PIXEL FOR APPEND-ON-SEND

       Outlook inserts this content when the message is
       actually sent.

       This does NOT use an OnMessageSend LaunchEvent.
    ===================================================== */

    await appendOnSendAsync(
      item.body,
      trackingContent,
      bodyFormat
    );


    /* =====================================================
       MARK THIS COMPOSE SESSION AS PREPARED
    ===================================================== */

    await sessionSetAsync(
      item.sessionData,
      TRACKING_SESSION_KEY,
      pixelUrl
    );


  } catch (_) {

    /*
     * FAIL OPEN.
     *
     * Any tracking failure is deliberately ignored.
     *
     * Tracking must remain completely independent from
     * Outlook's ability to send the message.
     */

  } finally {

    /*
     * Complete the recipient-change event.
     */

    event.completed();

  }

}


/* =========================================================
   OUTLOOK EVENT REGISTRATION

   IMPORTANT:
   v8 declares ONLY:

       OnMessageRecipientsChanged

   There is deliberately NO OnMessageSend association here.
========================================================= */

try {

  Office.actions.associate(
    "onMessageRecipientsChangedHandler",
    onMessageRecipientsChangedHandler
  );


  /*
   * DIAGNOSTIC STAGE 2
   *
   * If this appears, Office.actions.associate completed
   * successfully.
   */

  writeDiagnostic(
    "OUTLOOK_RECIPIENT_HANDLER_REGISTERED"
  );


} catch (error) {

  /*
   * Capture the actual Office exception rather than hiding
   * it behind a generic registration failure.
   */

  let errorMessage =
    "UNKNOWN_REGISTRATION_ERROR";


  try {

    if (error) {

      if (error.name) {

        errorMessage +=
          "_NAME_" +
          String(
            error.name
          );

      }


      if (error.message) {

        errorMessage +=
          "_MESSAGE_" +
          String(
            error.message
          );

      }


      if (
        !error.name &&
        !error.message
      ) {

        errorMessage +=
          "_" +
          String(
            error
          );

      }

    }

  } catch (_) {

    errorMessage =
      "ERROR_READING_REGISTRATION_EXCEPTION";

  }


  /*
   * Keep this reasonably short because Wix Subject fields
   * can be inconvenient to inspect when extremely long.
   */

  errorMessage =
    errorMessage
      .replace(
        /\s+/g,
        "_"
      )
      .slice(
        0,
        350
      );


  writeDiagnostic(
    "REGISTRATION_ERROR_" +
    errorMessage
  );

}
