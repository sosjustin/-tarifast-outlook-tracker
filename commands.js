const TRACKING_CREATE_URL =
  "https://www.tarifastops.com/_functions/emailTrackingCreate";

const TRACKING_SESSION_KEY =
  "tarifastEmailTrackingPixelUrl";


/* =========================================================
   DIAGNOSTIC HELPER

   Temporary diagnostic records written to Wix so we can
   determine exactly how far commands.js gets.

   This function is deliberately fail-open.
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
            stage,

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
    // Diagnostics must never interfere with Outlook.
  }
}


/* =========================================================
   DIAGNOSTIC STAGE 1

   If this appears in Wix, commands.js itself loaded and
   began executing.
========================================================= */

writeDiagnostic(
  "OUTLOOK_COMMANDS_JS_LOADED"
);


/* =========================================================
   SEND HANDLER

   This remains here for compatibility with the existing
   JavaScript file, but v8 does NOT declare OnMessageSend.

   Therefore Outlook should never invoke this function.
========================================================= */

function onMessageSendHandler(event) {
  event.completed({
    allowEvent: true
  });
}


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
        resolve(result.value);
      } else {
        reject(
          result.error ||
          new Error("Office getAsync failed")
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
        resolve(result.value);
      } else {
        reject(
          result.error ||
          new Error("body.getTypeAsync failed")
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

  const email = String(
    recipient.emailAddress || ""
  ).trim();

  if (!email) {
    return null;
  }

  return {
    email,

    name: String(
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
    Date.now().toString(36),
    Math.random()
      .toString(36)
      .slice(2, 12)
  ].join("-");
}


/* =========================================================
   CREATE TRACKING RECORD
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
    setTimeout(() => {
      if (controller) {
        controller.abort();
      }
    }, 5000);


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

            company: "",

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
========================================================= */

async function onMessageRecipientsChangedHandler(
  event
) {

  /*
   * DIAGNOSTIC STAGE 3
   *
   * This is deliberately before any Office mailbox calls.
   * If this appears, Outlook actually invoked our handler.
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
      typeof item.body
        .appendOnSendAsync !==
        "function" ||
      typeof item.body
        .getTypeAsync !==
        "function"
    ) {
      return;
    }


    /*
     * Prevent duplicate tracking records within the
     * same compose session.
     */

    const existingPixelUrl =
      await sessionGetAsync(
        item.sessionData,
        TRACKING_SESSION_KEY
      );


    if (existingPixelUrl) {
      return;
    }


    /*
     * Get current TO recipients.
     */

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


    /*
     * Get actual Outlook body format.
     */

    const bodyFormat =
      await bodyGetTypeAsync(
        item.body
      );


    /*
     * Subject is metadata only.
     */

    let subject = "";


    try {

      subject =
        await officeGetAsync(
          item.subject
        );

    } catch (_) {

      subject = "";

    }


    /*
     * Create tracking record in Wix.
     */

    const pixelUrl =
      await createTrackingRecord(
        recipient,
        subject
      );


    /*
     * Tracking pixels require an HTML message.
     */

    if (
      bodyFormat !==
      Office.CoercionType.Html
    ) {
      return;
    }


    const safePixelUrl =
      pixelUrl
        .replace(/&/g, "&amp;")
        .replace(/"/g, "&quot;");


    const trackingContent =
      '<img src="' +
      safePixelUrl +
      '" width="1" height="1" alt="" />';


    /*
     * Register pixel for insertion when Outlook sends
     * the message.
     */

    await appendOnSendAsync(
      item.body,
      trackingContent,
      bodyFormat
    );


    /*
     * Remember that this compose session is prepared.
     */

    await sessionSetAsync(
      item.sessionData,
      TRACKING_SESSION_KEY,
      pixelUrl
    );


  } catch (_) {

    /*
     * FAIL OPEN.
     *
     * Tracking failure never determines whether an
     * email can be sent.
     */

  } finally {

    event.completed();

  }
}


/* =========================================================
   OUTLOOK EVENT REGISTRATION
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
   * without throwing.
   */

  writeDiagnostic(
    "OUTLOOK_RECIPIENT_HANDLER_REGISTERED"
  );

} catch (_) {

  /*
   * If association itself fails, create a separate
   * diagnostic record.
   */

  writeDiagnostic(
    "OUTLOOK_RECIPIENT_HANDLER_REGISTRATION_FAILED"
  );

}


/*
 * Keep the existing send association in JavaScript.
 *
 * v8 does NOT contain an OnMessageSend LaunchEvent, so
 * Outlook should not invoke this.
 */

try {

  Office.actions.associate(
    "onMessageSendHandler",
    onMessageSendHandler
  );

} catch (_) {

  // Never interfere with Outlook.

}
