/* Tarifast Outlook Open Tracker
   Event-based OnMessageSend handler.
   Wix tracking endpoint is already live.
*/

const TRACKING_CREATE_URL =
  "https://www.tarifastops.com/_functions/emailTrackingCreate";

function getRecipientsAsync(field) {
  return new Promise((resolve) => {
    if (!field || typeof field.getAsync !== "function") {
      resolve([]);
      return;
    }

    field.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(Array.isArray(result.value) ? result.value : []);
      } else {
        resolve([]);
      }
    });
  });
}

function getSubjectAsync(item) {
  return new Promise((resolve) => {
    if (!item?.subject || typeof item.subject.getAsync !== "function") {
      resolve("");
      return;
    }

    item.subject.getAsync((result) => {
      resolve(
        result.status === Office.AsyncResultStatus.Succeeded
          ? String(result.value || "")
          : ""
      );
    });
  });
}

function getBodyHtmlAsync(item) {
  return new Promise((resolve, reject) => {
    item.body.getAsync(
      Office.CoercionType.Html,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(String(result.value || ""));
        } else {
          reject(new Error(result.error?.message || "Unable to read message body."));
        }
      }
    );
  });
}

function setBodyHtmlAsync(item, html) {
  return new Promise((resolve, reject) => {
    item.body.setAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(result.error?.message || "Unable to update message body."));
        }
      }
    );
  });
}

function normalizeAddress(recipient) {
  return String(
    recipient?.emailAddress ||
    recipient?.address ||
    ""
  ).trim().toLowerCase();
}

function normalizeName(recipient) {
  return String(
    recipient?.displayName ||
    recipient?.name ||
    ""
  ).trim();
}

function escapeHtmlAttribute(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

async function createTrackingRecord(recipient, subject) {
  const recipientEmail = normalizeAddress(recipient);

  if (!recipientEmail) {
    throw new Error("Recipient email address is unavailable.");
  }

  const response = await fetch(TRACKING_CREATE_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      recipientEmail,
      recipientName: normalizeName(recipient),
      company: "",
      subject: String(subject || ""),
      messageId: ""
    })
  });

  const data = await response.json();

  if (!response.ok || !data?.success || !data?.pixelUrl) {
    throw new Error(
      data?.error ||
      `Tracking endpoint returned HTTP ${response.status}.`
    );
  }

  return data;
}

async function onMessageSendHandler(event) {
  try {
    const item = Office.context.mailbox.item;

    const [toRecipients, subject, bodyHtml] = await Promise.all([
      getRecipientsAsync(item.to),
      getSubjectAsync(item),
      getBodyHtmlAsync(item)
    ]);

    /*
     * Phase 1:
     * Create one tracking record for the first TO recipient.
     * We deliberately do not create ambiguous duplicate pixels for
     * multi-recipient messages until per-recipient behavior is defined.
     */
    const primaryRecipient = toRecipients.find(r => normalizeAddress(r));

    if (!primaryRecipient) {
      event.completed({
        allowEvent: true
      });
      return;
    }

    const tracking = await createTrackingRecord(
      primaryRecipient,
      subject
    );

    const pixelUrl = escapeHtmlAttribute(tracking.pixelUrl);

    const trackingPixel =
      `<img src="${pixelUrl}" width="1" height="1" ` +
      `style="display:block;width:1px;height:1px;border:0;margin:0;padding:0;" ` +
      `alt="">`;

    /*
     * Put the pixel at the end of the existing HTML body.
     * Existing message content/signature is preserved.
     */
    await setBodyHtmlAsync(
      item,
      bodyHtml + trackingPixel
    );

    event.completed({
      allowEvent: true
    });

  } catch (error) {
    /*
     * SOFT-FAIL BY DESIGN:
     * Tracking must never prevent a Tarifast email from being sent.
     */
    event.completed({
      allowEvent: true
    });
  }
}

Office.actions.associate(
  "onMessageSendHandler",
  onMessageSendHandler
);
