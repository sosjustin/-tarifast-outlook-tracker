/* Tarifast Outlook Tracker - DIAGNOSTIC
   Tests:
   Outlook OnMessageSend -> recipient/subject -> Wix -> completion
   Does NOT modify the email body yet.
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
    if (!item || !item.subject || typeof item.subject.getAsync !== "function") {
      resolve("");
      return;
    }

    item.subject.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(String(result.value || ""));
      } else {
        resolve("");
      }
    });
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

async function createTrackingRecord(recipient, subject) {

  const recipientEmail = normalizeAddress(recipient);

  if (!recipientEmail) {
    throw new Error("Recipient email unavailable");
  }

  const controller = new AbortController();

  const timeout = setTimeout(() => {
    controller.abort();
  }, 8000);

  try {

    const response = await fetch(
      TRACKING_CREATE_URL,
      {
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
        }),

        signal: controller.signal
      }
    );

    if (!response.ok) {
      throw new Error(
        `Tracking endpoint HTTP ${response.status}`
      );
    }

    const data = await response.json();

    if (!data || !data.success) {
      throw new Error(
        data?.error ||
        "Tracking endpoint did not return success"
      );
    }

    return data;

  } finally {
    clearTimeout(timeout);
  }
}

async function onMessageSendHandler(event) {

  try {

    const item =
      Office.context.mailbox.item;

    const toRecipients =
      await getRecipientsAsync(item.to);

    const subject =
      await getSubjectAsync(item);

    const primaryRecipient =
      toRecipients.find(
        recipient =>
          normalizeAddress(recipient)
      );

    if (!primaryRecipient) {

      event.completed({
        allowEvent: true
      });

      return;
    }

    /*
     * DIAGNOSTIC:
     * Create Wix record only.
     *
     * No body.getAsync()
     * No body.setAsync()
     * No tracking pixel yet.
     */

    await createTrackingRecord(
      primaryRecipient,
      subject
    );

    event.completed({
      allowEvent: true
    });

  } catch (error) {

    /*
     * Tracking failure must NEVER
     * prevent the email from sending.
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
