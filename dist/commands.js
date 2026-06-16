/*
 * Outlook OnMessageSend handler.
 * This file is loaded by dist/commands.html.
 */

Office.onReady(() => {
  console.log("[DLP] Office.js ready in commands runtime");
});

/**
 * Main Outlook send-event handler.
 * The name must match the FunctionName in manifest.xml:
 *
 * <LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="SoftBlock"/>
 */
async function onMessageSendHandler(event) {
  console.log("[DLP] onMessageSendHandler started");

  try {
    const item = Office.context.mailbox.item;

    const subject = await getSubject(item);
    const body = await getBody(item);

    console.log("[DLP] Subject:", subject);
    console.log("[DLP] Body length:", body.length);
    console.log("[DLP] About to call fetch");

    const response = await fetch("https://9bq2s133w8.execute-api.us-east-1.amazonaws.com/analyze", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        subject: subject || "",
        body: (body || "").substring(0, 4000)
      })
    });

    console.log("[DLP] Fetch completed with status:", response.status);

    if (!response.ok) {
      console.error("[DLP] API returned non-OK status:", response.status);

      // Soft-fail open so mail is not blocked if API has an issue.
      event.completed({
        allowEvent: true
      });

      return;
    }

    const result = await response.json();

    console.log("[DLP] API result:", result);

    if (result && result.sensitive) {
      event.completed({
        allowEvent: false,
        errorMessage: result.reason || "Sensitive content detected. Please review this email before sending."
      });

      return;
    }

    event.completed({
      allowEvent: true
    });
  } catch (error) {
    console.error("[DLP] Error in onMessageSendHandler:", error);

    // Fail open during debugging.
    // Change this to allowEvent:false later if your policy requires blocking on errors.
    event.completed({
      allowEvent: true
    });
  }
}

/**
 * Gets email subject.
 */
function getSubject(item) {
  return new Promise((resolve, reject) => {
    try {
      item.subject.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value || "");
        } else {
          reject(result.error);
        }
      });
    } catch (error) {
      reject(error);
    }
  });
}

/**
 * Gets email body as plain text.
 */
function getBody(item) {
  return new Promise((resolve, reject) => {
    try {
      item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value || "");
        } else {
          reject(result.error);
        }
      });
    } catch (error) {
      reject(error);
    }
  });
}

/**
 * Required for event-based Outlook add-ins.
 * The string must match the FunctionName in manifest.xml.
 */
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
