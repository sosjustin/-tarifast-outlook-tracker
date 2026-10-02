/* ============================================================
   TARIFAST EMAIL TRACKER
   Outlook Event Runtime

   Event:
   OnMessageRecipientsChanged

   IMPORTANT:
   - No OnMessageSend handler.
   - No Smart Alerts send interception.
   - Tracking failure must never block sending.
   - Tracking pixel is registered with appendOnSendAsync.
============================================================ */


const TRACKING_CREATE_URL =
  "https://www.tarifastops.com/_functions/emailTrackingCreate";

const TRACKING_SESSION_KEY =
  "tarifastEmailTrackingPixelUrl";


/* ============================================================
   TEMPORARY DIAGNOSTIC

   Remove after production tracking is proven.
============================================================ */

function writeDiagnostic(stage) {

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
            "OUTLOOK JS DIAGNOSTIC",

          company:
            "Tarifast",

          subject:
            String(stage || "UNKNOWN"),

          messageId:
            "outlook-diagnostic-" +
            Date.now() +
            "-" +
            Math.random()
              .toString(36)
              .slice(2, 8)

        })

      }
    ).catch(function () {});

  } catch (_) {}

}


/* ============================================================
   PROVE COMMANDS.JS EXECUTED
============================================================ */

writeDiagnostic(
  "OUTLOOK_COMMANDS_JS_LOADED"
);


/* ============================================================
   OFFICE ASYNC HELPERS
============================================================ */

function getAsyncValue(target) {

  return new Promise(function (resolve, reject) {

    target.getAsync(function (result) {

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


function getBodyType(body) {

  return new Promise(function (resolve, reject) {

    body.getTypeAsync(function (result) {

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


function getSessionValue(
  sessionData,
  key
) {

  return new Promise(function (resolve) {

    sessionData.getAsync(
      key,
      function (result) {

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


function setSessionValue(
  sessionData,
  key,
  value
) {

  return new Promise(function (resolve) {

    sessionData.setAsync(
      key,
      value,
      function () {
        resolve();
      }
    );

  });

}


function appendTrackingPixel(
  body,
  html
) {

  return new Promise(function (resolve, reject) {

    body.appendOnSendAsync(
      html,
      {
        coercionType:
          Office.CoercionType.Html
      },

      function (result) {

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


/* ============================================================
   RECIPIENT
============================================================ */

function getFirstRecipient(recipients) {

  if (
    !Array.isArray(recipients) ||
    recipients.length === 0
  ) {

    return null;

  }


  const recipient =
    recipients[0];


  if (!recipient) {

    return null;

  }


  const email =
    String(
      recipient.emailAddress || ""
    ).trim();


  if (!email) {

    return null;

  }


  return {

    email: email,

    name:
      String(
        recipient.displayName || ""
      ).trim()

  };

}


/* ============================================================
   CLIENT MESSAGE ID
============================================================ */

function createClientMessageId() {

  return (
    "tarifast-" +
    Date.now().toString(36) +
    "-" +
    Math.random()
      .toString(36)
      .slice(2, 12)
  );

}


/* ============================================================
   CREATE WIX TRACKING RECORD
============================================================ */

function createTrackingRecord(
  recipient,
  subject
) {

  return fetch(
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
          createClientMessageId()

      })

    }
  )
  .then(function (response) {

    if (!response.ok) {

      throw new Error(
        "Tracking endpoint HTTP " +
        response.status
      );

    }

    return response.json();

  })
  .then(function (data) {

    if (
      !data ||
      data.success !== true ||
      !data.pixelUrl
    ) {

      throw new Error(
        "Tracking endpoint returned no pixel URL"
      );

    }


    return String(
      data.pixelUrl
    );

  });

}


/* ============================================================
   ESCAPE PIXEL URL
============================================================ */

function escapeHtmlAttribute(value) {

  return String(value)
    .replace(
      /&/g,
      "&amp;"
    )
    .replace(
      /"/g,
      "&quot;"
    )
    .replace(
      /</g,
      "&lt;"
    )
    .replace(
      />/g,
      "&gt;"
    );

}


/* ============================================================
   TRACK MESSAGE
============================================================ */

function prepareTrackingPixel() {

  const item =
    Office.context.mailbox.item;


  if (!item) {

    writeDiagnostic(
      "TRACKING_NO_ITEM"
    );

    return Promise.resolve();

  }


  if (
    !item.to ||
    !item.body
  ) {

    writeDiagnostic(
      "TRACKING_REQUIRED_APIS_MISSING"
    );

    return Promise.resolve();

  }


  /* ----------------------------------------------------------
     Get recipients.
  ---------------------------------------------------------- */

  return getAsyncValue(
    item.to
  )

  .then(function (recipients) {

    const recipient =
      getFirstRecipient(
        recipients
      );


    if (!recipient) {

      writeDiagnostic(
        "TRACKING_NO_TO_RECIPIENT"
      );

      return null;

    }


    return {

      recipient: recipient

    };

  })


  /* ----------------------------------------------------------
     Check body type.
  ---------------------------------------------------------- */

  .then(function (context) {

    if (!context) {

      return null;

    }


    return getBodyType(
      item.body
    )
    .then(function (bodyType) {

      context.bodyType =
        bodyType;

      return context;

    });

  })


  /* ----------------------------------------------------------
     Tracking pixel requires an HTML message.
  ---------------------------------------------------------- */

  .then(function (context) {

    if (!context) {

      return null;

    }


    if (
      context.bodyType !==
      Office.CoercionType.Html
    ) {

      writeDiagnostic(
        "TRACKING_BODY_NOT_HTML"
      );

      return null;

    }


    return context;

  })


  /* ----------------------------------------------------------
     Check session data if supported.

     If the message is already prepared, do nothing.
  ---------------------------------------------------------- */

  .then(function (context) {

    if (!context) {

      return null;

    }


    if (
      !item.sessionData ||
      typeof item.sessionData.getAsync !==
        "function"
    ) {

      context.sessionSupported =
        false;

      return context;

    }


    context.sessionSupported =
      true;


    return getSessionValue(
      item.sessionData,
      TRACKING_SESSION_KEY
    )
    .then(function (existingPixel) {

      if (existingPixel) {

        writeDiagnostic(
          "TRACKING_ALREADY_PREPARED"
        );

        return null;

      }


      return context;

    });

  })


  /* ----------------------------------------------------------
     Get subject.

     Subject failure does not stop tracking.
  ---------------------------------------------------------- */

  .then(function (context) {

    if (!context) {

      return null;

    }


    if (!item.subject) {

      context.subject = "";

      return context;

    }


    return getAsyncValue(
      item.subject
    )

    .then(function (subject) {

      context.subject =
        subject || "";

      return context;

    })

    .catch(function () {

      context.subject = "";

      return context;

    });

  })


  /* ----------------------------------------------------------
     Create the Wix tracking record.
  ---------------------------------------------------------- */

  .then(function (context) {

    if (!context) {

      return null;

    }


    return createTrackingRecord(
      context.recipient,
      context.subject
    )

    .then(function (pixelUrl) {

      context.pixelUrl =
        pixelUrl;

      writeDiagnostic(
        "TRACKING_RECORD_CREATED"
      );

      return context;

    });

  })


  /* ----------------------------------------------------------
     Register pixel for append-on-send.
  ---------------------------------------------------------- */

  .then(function (context) {

    if (!context) {

      return null;

    }


    if (
      typeof item.body.appendOnSendAsync !==
        "function"
    ) {

      writeDiagnostic(
        "APPEND_ON_SEND_NOT_AVAILABLE"
      );

      return null;

    }


    const safePixelUrl =
      escapeHtmlAttribute(
        context.pixelUrl
      );


    const pixelHtml =
      '<img src="' +
      safePixelUrl +
      '" width="1" height="1" alt="" ' +
      'style="display:block;width:1px;height:1px;border:0;" />';


    return appendTrackingPixel(
      item.body,
      pixelHtml
    )

    .then(function () {

      writeDiagnostic(
        "TRACKING_PIXEL_REGISTERED"
      );


      if (
        context.sessionSupported &&
        item.sessionData &&
        typeof item.sessionData.setAsync ===
          "function"
      ) {

        return setSessionValue(
          item.sessionData,
          TRACKING_SESSION_KEY,
          context.pixelUrl
        );

      }


      return null;

    })

    .then(function () {

      return context;

    });

  });


}


/* ============================================================
   ON MESSAGE RECIPIENTS CHANGED

   THIS IS THE FUNCTION NAMED BY THE v8 MANIFEST.
============================================================ */

function onMessageRecipientsChangedHandler(
  event
) {

  /*
   * This is the critical diagnostic.
   *
   * If this appears, Outlook has invoked the actual
   * OnMessageRecipientsChanged handler.
   */

  writeDiagnostic(
    "OUTLOOK_RECIPIENT_HANDLER_ENTERED"
  );


  try {

    prepareTrackingPixel()

      .catch(function (error) {

        let message =
          "TRACKING_FAILED";


        try {

          if (
            error &&
            error.message
          ) {

            message +=
              "_" +
              String(
                error.message
              );

          }

        } catch (_) {}


        writeDiagnostic(
          message
            .replace(
              /\s+/g,
              "_"
            )
            .slice(
              0,
              300
            )
        );

      })

      .then(function () {

        /*
         * Always complete the event.
         */

        try {

          event.completed();

        } catch (_) {}

      });


  } catch (error) {

    /*
     * Absolute fail-open protection.
     */

    writeDiagnostic(
      "TRACKING_HANDLER_EXCEPTION"
    );


    try {

      event.completed();

    } catch (_) {}

  }

}


/* ============================================================
   REGISTER EVENT HANDLER

   Microsoft requires the manifest FunctionName to be
   associated with the JavaScript function.
============================================================ */

try {

  Office.actions.associate(
    "onMessageRecipientsChangedHandler",
    onMessageRecipientsChangedHandler
  );


  writeDiagnostic(
    "OUTLOOK_RECIPIENT_HANDLER_REGISTERED"
  );


} catch (error) {

  let message =
    "OUTLOOK_RECIPIENT_HANDLER_REGISTRATION_FAILED";


  try {

    if (
      error &&
      error.message
    ) {

      message +=
        "_" +
        String(
          error.message
        );

    }

  } catch (_) {}


  writeDiagnostic(
    message
      .replace(
        /\s+/g,
        "_"
      )
      .slice(
        0,
        300
      )
  );

}
