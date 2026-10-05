/* ============================================================
   TARIFAST EMAIL TRACKER
   PIXEL-TEST-1 — HEADER + UNIQUE APPEND-ON-SEND PIXEL

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
   Receive unique trackingId
        ↓
   Stamp:
   X-Tarifast-Tracking-ID: <trackingId>
        ↓
   Register unique tracking pixel with appendOnSendAsync
        ↓
   Outlook sends normally
        ↓
   Pixel is appended only when the message is sent

   IMPORTANT
   ------------------------------------------------------------
   - NO OnMessageSend
   - NO Smart Alerts
   - NO send blocking
   - Exchange tracking rule remains DISABLED
   - Tracking failures must never prevent email use
   - Wix remains in diagnostic IMAGE_FETCHED mode
   - Uses handler name already proven with v11 manifest
============================================================ */

(function () {
  "use strict";

  const VERSION = "PIXEL-TEST-1";

  const TRACKING_CREATE_URL =
    "https://www.tarifastops.com/_functions/emailTrackingCreate";

  const TRACKING_OPEN_URL =
    "https://www.tarifastops.com/_functions/emailOpen";

  const TRACKING_HEADER =
    "X-Tarifast-Tracking-ID";


  /* ============================================================
     COMPOSE SESSION STATE
  ============================================================ */

  const composeId = createComposeId();

  let operationInProgress = false;

  let trackedRecipientKey = null;

  let currentTracking = null;

  /*
   * Tracks the exact trackingId currently registered
   * with appendOnSendAsync.
   *
   * This prevents repeated recipient-change events from
   * registering the same pixel over and over.
   */
  let registeredPixelTrackingId = null;


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
    } catch (_) {}

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
    } catch (_) {}
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


  function selectTrackingRecipient(recipients) {
    if (!recipients.length) {
      return null;
    }

    return recipients[0];
  }


  function buildRecipientKey(recipients) {
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
      typeof item.subject.getAsync !== "function"
    ) {
      return "";
    }

    try {
      const subject =
        await officeAsync(function (callback) {
          item.subject.getAsync(callback);
        });

      return String(subject || "");

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
      !result.trackingId
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
        composeId:
          composeId
      }
    );

    return result;
  }


  /* ============================================================
     STAMP TRACKING ID INTO INTERNET HEADER
  ============================================================ */

  async function setTrackingHeader(
    trackingId
  ) {
    const item =
      Office.context.mailbox.item;

    if (
      !item.internetHeaders ||
      typeof item.internetHeaders.setAsync !==
        "function"
    ) {
      throw new Error(
        "internetHeaders.setAsync is unavailable in this Outlook client"
      );
    }

    const cleanTrackingId =
      String(trackingId || "").trim();

    if (!cleanTrackingId) {
      throw new Error(
        "Cannot set tracking header without a trackingId"
      );
    }

    const headers = {};

    headers[TRACKING_HEADER] =
      cleanTrackingId;

    await new Promise(
      function (resolve, reject) {
        try {
          item.internetHeaders.setAsync(
            headers,
            function (result) {
              if (
                result &&
                result.status ===
                  Office.AsyncResultStatus.Succeeded
              ) {
                resolve();
                return;
              }

              const message =
                result &&
                result.error &&
                result.error.message
                  ? result.error.message
                  : "Unknown internetHeaders.setAsync error";

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
      "TRACKING HEADER SET",
      {
        header:
          TRACKING_HEADER,
        trackingId:
          cleanTrackingId
      }
    );
  }


  /* ============================================================
     VERIFY HEADER IN CURRENT COMPOSE ITEM
  ============================================================ */

  async function verifyTrackingHeader(
    expectedTrackingId
  ) {
    const item =
      Office.context.mailbox.item;

    if (
      !item.internetHeaders ||
      typeof item.internetHeaders.getAsync !==
        "function"
    ) {
      log(
        "HEADER VERIFY UNAVAILABLE"
      );

      return;
    }

    try {
      const result =
        await officeAsync(
          function (callback) {
            item.internetHeaders.getAsync(
              [TRACKING_HEADER],
              callback
            );
          }
        );

      log(
        "TRACKING HEADER VERIFIED",
        {
          expected:
            expectedTrackingId,
          actual:
            result &&
            result[TRACKING_HEADER]
              ? result[TRACKING_HEADER]
              : ""
        }
      );

    } catch (error) {
      /*
       * Verification is diagnostic only.
       * Header set success remains authoritative.
       */
      log(
        "HEADER VERIFY ERROR",
        String(error)
      );
    }
  }


  /* ============================================================
     BUILD UNIQUE TRACKING PIXEL
  ============================================================ */

  function buildTrackingPixelHtml(
    trackingId
  ) {
    const cleanTrackingId =
      String(trackingId || "").trim();

    if (!cleanTrackingId) {
      throw new Error(
        "Cannot build tracking pixel without a trackingId"
      );
    }

    const pixelUrl =
      TRACKING_OPEN_URL +
      "?t=" +
      encodeURIComponent(
        cleanTrackingId
      );

    /*
     * Intentionally tiny and visually inert.
     *
     * Wix is currently diagnostic-only:
     * a request should record IMAGE_FETCHED telemetry,
     * NOT blindly increment a human-open count.
     */
    return (
      '<img src="' +
      pixelUrl +
      '" ' +
      'width="1" ' +
      'height="1" ' +
      'style="width:1px;height:1px;border:0;display:block;opacity:0;" ' +
      'alt="">'
    );
  }


  /* ============================================================
     REGISTER PIXEL FOR APPEND ON SEND
  ============================================================ */

  async function registerPixelOnSend(
    trackingId
  ) {
    const item =
      Office.context.mailbox.item;

    const cleanTrackingId =
      String(trackingId || "").trim();

    if (!cleanTrackingId) {
      throw new Error(
        "Cannot register tracking pixel without a trackingId"
      );
    }

    /*
     * If this exact trackingId is already registered,
     * do not register it again.
     */
    if (
      registeredPixelTrackingId ===
        cleanTrackingId
    ) {
      log(
        "PIXEL ALREADY REGISTERED — SKIPPING",
        {
          trackingId:
            cleanTrackingId
        }
      );

      return;
    }

    if (
      !item.body ||
      typeof item.body.getTypeAsync !==
        "function"
    ) {
      throw new Error(
        "body.getTypeAsync is unavailable in this Outlook client"
      );
    }

    if (
      typeof item.body.appendOnSendAsync !==
        "function"
    ) {
      throw new Error(
        "body.appendOnSendAsync is unavailable in this Outlook client"
      );
    }

    /*
     * Microsoft recommends reading the current body format
     * first, then passing that format to appendOnSendAsync.
     */
    const bodyType =
      await officeAsync(
        function (callback) {
          item.body.getTypeAsync(
            callback
          );
        }
      );

    let pixelContent;

    /*
     * HTML tracking requires an <img>.
     *
     * If Outlook somehow gives us a plain-text body,
     * do NOT attempt to force HTML into it.
     * Fail the tracking registration open instead of
     * affecting the user's message.
     */
    if (
      bodyType !==
        Office.CoercionType.Html
    ) {
      throw new Error(
        "Tracking pixel not registered because message body is not HTML"
      );
    }

    pixelContent =
      buildTrackingPixelHtml(
        cleanTrackingId
      );

    await new Promise(
      function (resolve, reject) {
        try {
          item.body.appendOnSendAsync(
            pixelContent,
            {
              coercionType:
                bodyType
            },
            function (result) {
              if (
                result &&
                result.status ===
                  Office.AsyncResultStatus.Succeeded
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

    /*
     * Only mark the trackingId registered AFTER
     * Outlook confirms appendOnSendAsync succeeded.
     */
    registeredPixelTrackingId =
      cleanTrackingId;

    log(
      "PIXEL REGISTERED FOR APPEND ON SEND",
      {
        trackingId:
          cleanTrackingId,

        bodyType:
          bodyType,

        pixelUrl:
          TRACKING_OPEN_URL +
          "?t=" +
          encodeURIComponent(
            cleanTrackingId
          )
      }
    );
  }


  /* ============================================================
     MAIN RECIPIENT-CHANGE HANDLER

     Function name intentionally remains synchronized
     with the proven v11 manifest.
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

        return;
      }

      const recipientKey =
        buildRecipientKey(
          recipients
        );

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

      const tracking =
        await createTrackingRecord(
          trackingRecipient,
          subject
        );

      /*
       * Keep the working custom header.
       */
      await setTrackingHeader(
        tracking.trackingId
      );

      await verifyTrackingHeader(
        tracking.trackingId
      );

      /*
       * Register the SAME unique trackingId as a
       * tracking pixel that Outlook will append
       * only when the message is sent.
       */
      await registerPixelOnSend(
        tracking.trackingId
      );

      currentTracking =
        tracking;

      trackedRecipientKey =
        recipientKey;

      log(
        "TRACKING PREPARED SUCCESSFULLY",
        {
          trackingId:
            tracking.trackingId,

          composeId:
            composeId,

          recipientKey:
            recipientKey,

          header:
            TRACKING_HEADER,

          pixelRegistered:
            true
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
