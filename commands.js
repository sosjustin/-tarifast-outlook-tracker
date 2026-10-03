/* ============================================================
   TARIFAST EMAIL TRACKER
   PRODUCTION-2 — COMPOSE SESSION ID + APPEND ON SEND

   ARCHITECTURE
   ------------------------------------------------------------
   Outlook OnMessageRecipientsChanged
        ↓
   Establish one composeId for this compose runtime
        ↓
   Read current To / Cc / Bcc recipients
        ↓
   Select trackable recipient
        ↓
   Create / reuse tracking record in Tarifast / Wix
        ↓
   Receive unique pixel URL
        ↓
   Register invisible pixel with appendOnSendAsync
        ↓
   Outlook sends normally

   IMPORTANT
   ------------------------------------------------------------
   - NO OnMessageSend
   - NO send blocking
   - Tracking failures must never prevent email use
   - Uses the handler name already proven with v11 manifest
   - composeId remains stable for this compose runtime
============================================================ */

(function () {
  "use strict";

  const VERSION = "PRODUCTION-2";

  const TRACKING_CREATE_URL =
    "https://www.tarifastops.com/_functions/emailTrackingCreate";


  /* ============================================================
     COMPOSE SESSION STATE
  ============================================================ */

  /*
   * One stable ID for this JavaScript runtime / compose session.
   *
   * Wix will use this in the next step to ensure repeated
   * recipient-change events reuse one EmailTracking record
   * instead of inserting duplicate/orphan records.
   */
  const composeId = createComposeId();

  /*
   * Prevent overlapping recipient-change processing.
   */
  let operationInProgress = false;

  /*
   * Remember the recipient set already armed during this runtime.
   */
  let trackedRecipientKey = null;

  /*
   * Remember the tracking record returned by Wix.
   *
   * Once the Wix endpoint becomes composeId-aware, repeated
   * create requests for this composeId will return the same
   * tracking record.
   */
  let currentTracking = null;


  /* ============================================================
     COMPOSE ID
  ============================================================ */

  function createComposeId() {
    try {
      if (
        typeof crypto !== "undefined" &&
        typeof crypto.randomUUID === "function"
      ) {
        return crypto.randomUUID();
      }
    } catch (_) {
      // Fall through to compatibility generator.
    }

    return (
      "tf-" +
      Date.now().toString(36) +
      "-" +
      Math.random().toString(36).slice(2, 12) +
      "-" +
      Math.random().toString(36).slice(2, 12)
    );
  }


  /* ============================================================
     DIAGNOSTICS
  ============================================================ */

  function log(message, data) {
    try {
      if (typeof data !== "undefined") {
        console.log(
          "TARIFAST EMAIL TRACKER | " +
            VERSION +
            " | " +
            new Date().toISOString() +
            " | " +
            message,
          data
        );
      } else {
        console.log(
          "TARIFAST EMAIL TRACKER | " +
            VERSION +
            " | " +
            new Date().toISOString() +
            " | " +
            message
        );
      }
    } catch (_) {
      // Logging must never interfere with Outlook.
    }
  }


  /* ============================================================
     SAFE EVENT COMPLETION
  ============================================================ */

  function completeEvent(event) {
    try {
      if (
        event &&
        typeof event.completed === "function"
      ) {
        event.completed();
      }
    } catch (error) {
      log(
        "event.completed ERROR",
        String(error)
      );
    }
  }


  /* ============================================================
     OFFICE ASYNC WRAPPER
  ============================================================ */

  function officeAsync(invoker) {
    return new Promise(function (resolve, reject) {
      try {
        invoker(function (result) {
          if (
            result &&
            result.status ===
              Office.AsyncResultStatus.Succeeded
          ) {
            resolve(result.value);
            return;
          }

          const message =
            result &&
            result.error &&
            result.error.message
              ? result.error.message
              : "Unknown Office.js error";

          reject(new Error(message));
        });

      } catch (error) {
        reject(error);
      }
    });
  }


  /* ============================================================
     RECIPIENT HELPERS
  ============================================================ */

  async function getRecipients(recipientField) {
    if (
      !recipientField ||
      typeof recipientField.getAsync !== "function"
    ) {
      return [];
    }

    try {
      const recipients =
        await officeAsync(function (callback) {
          recipientField.getAsync(callback);
        });

      return Array.isArray(recipients)
        ? recipients
        : [];

    } catch (error) {
      log(
        "RECIPIENT READ ERROR",
        String(error)
      );

      return [];
    }
  }


  async function getAllRecipients() {
    const item =
      Office.context.mailbox.item;

    const results =
      await Promise.all([
        getRecipients(item.to),
        getRecipients(item.cc),
        getRecipients(item.bcc)
      ]);

    return []
      .concat(results[0])
      .concat(results[1])
      .concat(results[2]);
  }


  function normalizeEmail(value) {
    return String(value || "")
      .trim()
      .toLowerCase();
  }


  function uniqueRecipients(recipients) {
    const seen =
      Object.create(null);

    return recipients.filter(
      function (recipient) {
        const email =
          normalizeEmail(
            recipient.emailAddress
          );

        if (
          !email ||
          seen[email]
        ) {
          return false;
        }

        seen[email] = true;

        return true;
      }
    );
  }


  /*
   * Current production behavior:
   * one tracking record / pixel follows the first
   * resolved recipient.
   */
  function selectTrackingRecipient(
    recipients
  ) {
    if (!recipients.length) {
      return null;
    }

    return recipients[0];
  }


  function buildRecipientKey(
    recipients
  ) {
    return recipients
      .map(function (recipient) {
        return normalizeEmail(
          recipient.emailAddress
        );
      })
      .filter(Boolean)
      .sort()
      .join("|");
  }


  /* ============================================================
     SUBJECT
  ============================================================ */

  async function getSubject() {
    const item =
      Office.context.mailbox.item;

    if (
      !item.subject ||
      typeof item.subject.getAsync !==
        "function"
    ) {
      return "";
    }

    try {
      const subject =
        await officeAsync(
          function (callback) {
            item.subject.getAsync(
              callback
            );
          }
        );

      return String(
        subject || ""
      );

    } catch (error) {
      log(
        "SUBJECT READ ERROR",
        String(error)
      );

      return "";
    }
  }


  /* ============================================================
     CREATE / REUSE TRACKING RECORD
  ============================================================ */

  async function createTrackingRecord(
    recipient,
    subject
  ) {
    const payload = {
      composeId: composeId,

      recipientEmail:
        normalizeEmail(
          recipient.emailAddress
        ),

      recipientName:
        String(
          recipient.displayName || ""
        ).trim(),

      company: "",

      subject:
        String(
          subject || ""
        ).trim(),

      messageId: ""
    };

    log(
      "CREATING OR REUSING TRACKING RECORD",
      payload
    );

    const response =
      await fetch(
        TRACKING_CREATE_URL,
        {
          method: "POST",

          headers: {
            "Content-Type":
              "application/json"
          },

          body:
            JSON.stringify(payload)
        }
      );

    if (!response.ok) {
      throw new Error(
        "Tracking endpoint HTTP " +
          response.status +
          " " +
          response.statusText
      );
    }

    const result =
      await response.json();

    if (
      !result ||
      result.success !== true ||
      !result.pixelUrl
    ) {
      throw new Error(
        "Tracking endpoint returned an invalid response"
      );
    }

    log(
      "TRACKING RECORD READY",
      {
        id: result.id,
        trackingId:
          result.trackingId,
        pixelUrl:
          result.pixelUrl,
        composeId:
          composeId
      }
    );

    return result;
  }


  /* ============================================================
     PIXEL HTML
  ============================================================ */

  function buildPixelHtml(
    pixelUrl
  ) {
    const safeUrl =
      String(pixelUrl || "")
        .replace(/&/g, "&amp;")
        .replace(/"/g, "&quot;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");

    return (
      '<img src="' +
      safeUrl +
      '" width="1" height="1" ' +
      'style="width:1px;height:1px;border:0;' +
      'display:block;opacity:0;" alt="">'
    );
  }


  /* ============================================================
     APPEND PIXEL ON SEND
  ============================================================ */

  async function registerPixelOnSend(
    pixelUrl
  ) {
    const item =
      Office.context.mailbox.item;

    if (
      !item.body ||
      typeof item.body.appendOnSendAsync !==
        "function"
    ) {
      throw new Error(
        "appendOnSendAsync is unavailable in this Outlook client"
      );
    }

    const bodyType =
      await officeAsync(
        function (callback) {
          item.body.getTypeAsync(
            callback
          );
        }
      );

    /*
     * HTML pixels cannot be inserted into
     * a plain-text message body.
     */
    if (
      bodyType !==
      Office.CoercionType.Html
    ) {
      throw new Error(
        "Message body is not HTML; tracking pixel was not registered"
      );
    }

    const pixelHtml =
      buildPixelHtml(
        pixelUrl
      );

    await new Promise(
      function (resolve, reject) {
        try {
          item.body.appendOnSendAsync(
            pixelHtml,
            {
              coercionType:
                Office.CoercionType.Html
            },
            function (result) {
              if (
                result &&
                result.status ===
                  Office.AsyncResultStatus
                    .Succeeded
              ) {
                resolve();
                return;
              }

              const message =
                result &&
                result.error &&
                result.error.message
                  ? result.error.message
                  : "Unknown appendOnSendAsync error";

              reject(
                new Error(message)
              );
            }
          );

        } catch (error) {
          reject(error);
        }
      }
    );

    log(
      "TRACKING PIXEL REGISTERED FOR SEND"
    );
  }


  /* ============================================================
     MAIN RECIPIENT-CHANGE HANDLER

     IMPORTANT:
     This function name remains synchronized with
     the proven v11 manifest:
     tarifastV10RecipientsChanged
  ============================================================ */

  async function tarifastV10RecipientsChanged(
    event
  ) {
    log(
      "RECIPIENT CHANGE FIRED"
    );

    if (operationInProgress) {
      log(
        "TRACKING OPERATION ALREADY IN PROGRESS"
      );

      completeEvent(event);
      return;
    }

    operationInProgress = true;

    try {
      const recipients =
        uniqueRecipients(
          await getAllRecipients()
        );

      if (!recipients.length) {
        log(
          "NO RECIPIENTS — NOTHING TO TRACK"
        );

        trackedRecipientKey = null;

        /*
         * Do NOT discard composeId.
         *
         * This is still the same compose session,
         * even if all recipients are temporarily
         * removed.
         */
        return;
      }

      const recipientKey =
        buildRecipientKey(
          recipients
        );

      /*
       * Same recipient set during this runtime:
       * nothing needs to be recreated or rearmed.
       */
      if (
        trackedRecipientKey &&
        trackedRecipientKey ===
          recipientKey
      ) {
        log(
          "RECIPIENT SET ALREADY TRACKED — SKIPPING"
        );

        return;
      }

      const trackingRecipient =
        selectTrackingRecipient(
          recipients
        );

      if (!trackingRecipient) {
        log(
          "NO TRACKABLE RECIPIENT FOUND"
        );

        return;
      }

      log(
        "TRACKING RECIPIENT",
        {
          email:
            trackingRecipient
              .emailAddress || "",

          name:
            trackingRecipient
              .displayName || "",

          composeId:
            composeId
        }
      );

      const subject =
        await getSubject();

      /*
       * After the Wix endpoint is upgraded,
       * this call becomes an UPSERT:
       *
       * same composeId = same EmailTracking record.
       */
      const tracking =
        await createTrackingRecord(
          trackingRecipient,
          subject
        );

      /*
       * Avoid registering the same tracking pixel
       * more than once when Wix returns the existing
       * tracking record for this compose session.
       *
       * If this is the first successful registration,
       * arm it now.
       */
      const alreadyRegistered =
        currentTracking &&
        currentTracking.trackingId &&
        currentTracking.trackingId ===
          tracking.trackingId;

      if (!alreadyRegistered) {
        await registerPixelOnSend(
          tracking.pixelUrl
        );
      } else {
        log(
          "TRACKING PIXEL ALREADY REGISTERED — REUSING"
        );
      }

      currentTracking =
        tracking;

      trackedRecipientKey =
        recipientKey;

      log(
        "TRACKING ARMED SUCCESSFULLY",
        {
          trackingId:
            tracking.trackingId,

          composeId:
            composeId,

          recipientKey:
            recipientKey
        }
      );

    } catch (error) {
      /*
       * FAIL OPEN.
       *
       * Tracking problems must never interfere
       * with normal Outlook compose/send behavior.
       */
      log(
        "TRACKING FAILED OPEN",
        error &&
        error.message
          ? error.message
          : String(error)
      );

    } finally {
      operationInProgress = false;

      /*
       * Always release Outlook's event.
       */
      completeEvent(event);
    }
  }


  /* ============================================================
     EVENT ASSOCIATION
  ============================================================ */

  try {
    Office.actions.associate(
      "tarifastV10RecipientsChanged",
      tarifastV10RecipientsChanged
    );

    log(
      "RECIPIENT HANDLER ASSOCIATION RETURNED"
    );

  } catch (error) {
    log(
      "RECIPIENT HANDLER ASSOCIATION ERROR",
      String(error)
    );
  }


  /* ============================================================
     NORMAL WEBVIEW INITIALIZATION DIAGNOSTIC
  ============================================================ */

  Office.initialize =
    function () {
      log(
        "OFFICE INITIALIZE CALLBACK"
      );
    };


  log(
    "SCRIPT EXECUTING",
    {
      composeId:
        composeId
    }
  );

})();
