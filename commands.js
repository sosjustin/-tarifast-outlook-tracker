const TRACKING_CREATE_URL =
  "https://www.tarifastops.com/_functions/emailTrackingCreate";

const TRACKING_SESSION_KEY =
  "tarifastEmailTrackingPixelUrl";


/* =========================================================
   SEND HANDLER

   IMPORTANT:
   This preserves the proven fail-open behavior.
   Tracking never determines whether the email can send.
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


function sessionGetAsync(sessionData, key) {
  return new Promise((resolve) => {
    sessionData.getAsync(
      key,
      (result) => {
        if (
          result.status ===
          Office.AsyncResultStatus.Succeeded
        ) {
          resolve(result.value || "");
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

          signal: controller
            ? controller.signal
            : undefined
        }
      );


    if (!response.ok) {
      throw new Error(
        `Tracking endpoint returned HTTP ${response.status}`
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

   Prepares the tracking pixel while the message is being
   composed.

   FAILURE HERE NEVER PREVENTS SENDING.
========================================================= */
async function onMessageRecipientsChangedHandler(
  event
) {

  try {

    /*
     * TEMPORARY DIAGNOSTIC
     * Proves Outlook actually launched this event handler.
     */

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
              "OUTLOOK EVENT DIAGNOSTIC",
            company:
              "Tarifast",
            subject:
              "OUTLOOK_RECIPIENT_EVENT",
            messageId:
              "recipient-event-" +
              Date.now()
          })
        }
      );
    } catch (_) {
      // Diagnostic must never interfere with Outlook.
    }


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
     * Prevent duplicate tracking records within the same
     * compose session.
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
     * Get the actual Outlook body format.
     *
     * Microsoft recommends passing this value directly
     * into appendOnSendAsync.
     */

    const bodyFormat =
      await bodyGetTypeAsync(
        item.body
      );


    /*
     * Subject is metadata only.
     * Tracking can still continue if Outlook cannot return it.
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
     * Create the tracking record in Wix.
     */

    const pixelUrl =
      await createTrackingRecord(
        recipient,
        subject
      );


    /*
     * Build content appropriate for the Outlook body format.
     */

    let trackingContent;


    if (
      bodyFormat ===
      Office.CoercionType.Html
    ) {

      const safePixelUrl =
        pixelUrl
          .replace(/&/g, "&amp;")
          .replace(/"/g, "&quot;");


      trackingContent =
        '<img src="' +
        safePixelUrl +
        '" width="1" height="1" alt="" />';

    } else {

      /*
       * A tracking pixel cannot operate inside a plain-text
       * message. Exit without affecting the message.
       */

      return;

    }


    /*
     * Register the pixel with Outlook.
     *
     * Outlook inserts it when the message is actually sent.
     */

    await appendOnSendAsync(
      item.body,
      trackingContent,
      bodyFormat
    );


    /*
     * Remember that this compose session has already been
     * prepared.
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
     * Any tracking error is deliberately ignored.
     * Sending remains independent from tracking.
     */

  } finally {

    event.completed();

  }
}


/* =========================================================
   OUTLOOK EVENT REGISTRATION
========================================================= */

Office.actions.associate(
  "onMessageRecipientsChangedHandler",
  onMessageRecipientsChangedHandler
);


Office.actions.associate(
  "onMessageSendHandler",
  onMessageSendHandler
);
