/*
 * Outlook Smart Alerts DLP handler.
 *
 * Flow:
 * 1. User clicks Send.
 * 2. OnMessageSend scans subject/body.
 * 3. If sensitive:
 *    - Show Smart Alert.
 *    - User can choose:
 *      a) Encrypt and Send -> function command adds [ENCRYPT] and sends.
 *      b) Send Anyway -> Outlook sends as-is.
 *      c) Don't Send -> message stays open.
 */

Office.onReady(() => {
  console.log("[DLP] Office.js ready");
});

const API_URL = "https://9bq2s133w8.execute-api.us-east-1.amazonaws.com/analyze";
const ENCRYPT_PREFIX = "[ENCRYPT]";

/**
 * Main OnMessageSend handler.
 * This must match manifest:
 * FunctionName="onMessageSendHandler"
 */
async function onMessageSendHandler(event) {
  console.log("[DLP] onMessageSendHandler started");

  try {
    const item = Office.context.mailbox.item;

    const subject = await getSubject(item);
    console.log("[DLP] Subject:", subject);

    // If user already selected Encrypt and Send, allow send.
    if (hasEncryptPrefix(subject)) {
      console.log("[DLP] Subject already encrypted. Allowing send.");
      event.completed({ allowEvent: true });
      return;
    }

    const body = await getBody(item);
    console.log("[DLP] Body length:", body.length);

    const result = await analyzeEmail(subject, body);
    console.log("[DLP] API result:", result);

    if (result && result.sensitive) {
      const reason = result.reason || "Sensitive content detected.";

      event.completed({
        allowEvent: false,

        // Message shown in Outlook Smart Alert.
        errorMessage:
          `${reason}\n\nChoose "Encrypt and Send" to prepend [ENCRYPT] to the subject, or choose "Send Anyway" to send without encryption.`,

        // This shows Outlook's built-in "Send Anyway" option.
        sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser,

        // This adds a Smart Alert action button.
        // This ID must match the manifest Control id.
        commandId: "encryptAndSendCommand",

        // Label for the action button. Max 20 chars.
        cancelLabel: "Encrypt and Send"
      });

      return;
    }

    console.log("[DLP] No sensitive content. Allowing send.");
    event.completed({ allowEvent: true });
  } catch (error) {
    console.error("[DLP] Error in onMessageSendHandler:", error);

    // During testing, fail open.
    // For strict DLP, change allowEvent to false.
    event.completed({
      allowEvent: true
    });
  }
}

/**
 * Smart Alert action button handler.
 * This runs when user clicks "Encrypt and Send".
 *
 * This requires:
 * - Mailbox requirement set 1.15
 * - ReadWriteMailbox permission
 * - Function command in manifest with id="encryptAndSendCommand"
 */
async function encryptAndSendHandler(event) {
  console.log("[DLP] encryptAndSendHandler started");

  try {
    const item = Office.context.mailbox.item;

    const currentSubject = await getSubject(item);
    const newSubject = hasEncryptPrefix(currentSubject)
      ? currentSubject
      : `${ENCRYPT_PREFIX} ${currentSubject || ""}`;

    console.log("[DLP] Setting encrypted subject:", newSubject);

    await setSubject(item, newSubject);

    console.log("[DLP] Subject updated. Sending message.");

    await sendMessage(item);

    // Function command completion.
    if (event && typeof event.completed === "function") {
      event.completed();
    }
  } catch (error) {
    console.error("[DLP] Encrypt and Send failed:", error);

    const item = Office.context.mailbox.item;
    addNotification(item, "Failed to encrypt and send. Please try again.");

    if (event && typeof event.completed === "function") {
      event.completed();
    }
  }
}

async function analyzeEmail(subject, body) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 5000);

  try {
    const response = await fetch(API_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        subject: subject || "",
        body: (body || "").substring(0, 4000)
      }),
      signal: controller.signal
    });

    clearTimeout(timeout);

    if (!response.ok) {
      throw new Error(`DLP API returned ${response.status}`);
    }

    return await response.json();
  } catch (error) {
    clearTimeout(timeout);
    throw error;
  }
}

function getSubject(item) {
  return new Promise((resolve, reject) => {
    item.subject.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || "");
      } else {
        reject(result.error);
      }
    });
  });
}

function setSubject(item, subject) {
  return new Promise((resolve, reject) => {
    item.subject.setAsync(subject, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(result.error);
      }
    });
  });
}

function getBody(item) {
  return new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || "");
      } else {
        reject(result.error);
      }
    });
  });
}

function sendMessage(item) {
  return new Promise((resolve, reject) => {
    if (!item.sendAsync) {
      reject(new Error("sendAsync is not supported in this Outlook client."));
      return;
    }

    item.sendAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(result.error);
      }
    });
  });
}

function hasEncryptPrefix(subject) {
  return (subject || "").trim().toUpperCase().startsWith(ENCRYPT_PREFIX);
}

function addNotification(item, message) {
  try {
    item.notificationMessages.addAsync("dlp-warning", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message,
      icon: "icon16",
      persistent: false
    });
  } catch (error) {
    console.error("[DLP] Failed to add notification:", error);
  }
}

/**
 * Required registrations.
 */
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("encryptAndSendHandler", encryptAndSendHandler);
