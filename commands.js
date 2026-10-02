const TRACKING_CREATE_URL =
  "https://www.tarifastops.com/_functions/emailTrackingCreate";

const TRACKING_SESSION_KEY = "tarifastEmailTrackingPixelUrl";


/* =========================================================
   SEND HANDLER

   IMPORTANT:
   This intentionally preserves the working fail-open behavior.
   Email sending is never dependent on tracking succeeding.
========================================================= */

function onMessageSendHandler(event) {
  event.completed({
    allowEvent: true
  });
}


/* =========================================================
   SMALL PROMISE WRAPPERS FOR OFFICE.JS
========================================================= */

function officeGetAsync(target) {
  return new Promise((resolve, reject) => {
    target.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error || new Error("Office getAsync failed"));
      }
    });
  });
}


function sessionGetAsync(sessionData, key) {
  return new Promise((resolve) => {
    sessionData.getAsync(key, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || "");
      } else {
        resolve("");
      }
    });
  });
}


function sessionSetAsync(sessionData, key, value) {
  return new Promise((resolve) => {
    sessionData.setAsync(key, value, () => resolve());
  });
}


function appendOnSendAsync(body, html) {
  return new Promise((resolve, reject) => {
    body.appendOnSendAsync(
      html,
      {
        coercionType: Office.CoercionType.Html
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(
            result.error ||
            new Error("appendOnSendAsync failed")
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
  if (!Array.isArray(recipients) || recipients.length === 0) {
    return null;
  }

  const recipient = recipients[0] || {};

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

   Outlook does not necessarily have its final internet
   Message-ID while the message is still being composed.
========================================================= */

function makeClientMessageId() {
  return [
    "tarifast",
    Date.now().toString(36),
    Math.random().toString(36).slice(2, 12)
  ].join("-");
}


/* =========================================================
   CREATE TRACKING RECORD IN WIX
========================================================= */

async function createTrackingRecord(recipient, subject) {

  const controller =
    typeof AbortController !== "undefined"
      ? new AbortController()
      : null;

  const timeoutId = setTimeout(() => {
    if (controller) {
      controller.abort();
    }
  }, 5000);

  try {

    const response = await fetch(
      TRACKING_CREATE_URL,
      {
        method: "POST",

        headers: {
          "Content-Type": "application/json"
        },

        body: JSON.stringify({
          recipientEmail: recipient.email,
          recipientName: recipient.name,
          company: "",
          subject: String(subject || "").trim(),
          messageId: makeClientMessageId()
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


    const data = await response.json();


    if (
      !data ||
      data.success !== true ||
      !data.pixelUrl
    ) {
      throw new Error(
        "Tracking endpoint did not return a pixel URL"
      );
    }


    return String(data.pixelUrl);

  } finally {

    clearTimeout(timeoutId);

  }
}


/* =========================================================
   RECIPIENT CHANGE EVENT

   This prepares the tracking pixel while the email is being
   composed.

   It does NOT control whether the email can be sent.

   Any failure simply exits and completes the event.
========================================================= */

async function onMessageRecipientsChangedHandler(event) {

  try {

    const item =
      Office.context.mailbox.item;


    if (
      !item ||
      !item.to ||
      !item.subject ||
      !item.body ||
      !item.sessionData ||
      typeof item.body.appendOnSendAsync !== "function"
    ) {
      return;
    }


    /*
     * Do not create multiple tracking records for the same
     * compose session if Outlook fires this event repeatedly.
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
      await officeGetAsync(item.to);


    const recipient =
      firstRecipient(recipients);


    if (!recipient) {
      return;
    }


    /*
     * Tracking pixels require an HTML body.
     */

    const bodyType =
      await officeGetAsync(item.body);


    if (
      bodyType !== Office.CoercionType.Html
    ) {
      return;
    }


    /*
     * Subject is useful metadata but is not required for
     * tracking. If Outlook cannot return it, continue anyway.
     */

    let subject = "";


    try {

      subject =
        await officeGetAsync(item.subject);

    } catch (_) {

      subject = "";

    }


    /*
     * Create the tracking record.
     */

    const pixelUrl =
      await createTrackingRecord(
        recipient,
        subject
      );


    /*
     * Build an invisible 1x1 tracking image.
     */

    const safePixelUrl =
      pixelUrl
        .replace(/&/g, "&amp;")
        .replace(/"/g, "&quot;");


    const pixelHtml =
      '<img src="' +
      safePixelUrl +
      '" width="1" height="1" alt="" ' +
      'style="width:1px;height:1px;border:0;margin:0;padding:0;" />';


    /*
     * Tell Outlook to append the pixel when the message
     * actually sends.
     */

    await appendOnSendAsync(
      item.body,
      pixelHtml
    );


    /*
     * Mark this compose session as prepared so repeated
     * recipient-change events do not create duplicate records.
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
     * Tracking failure must NEVER interfere with composing
     * or sending an email.
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
