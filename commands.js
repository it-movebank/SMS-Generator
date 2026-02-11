/* commands.js
 * Outlook add-in command function for ClickSend SMS verification.
 * Inserts a 6-digit code into the SUBJECT (because ClickSend uses subject as SMS text).
 *
 * Notes:
 * - Function commands must accept a single `event` parameter and call event.completed() when done. [1](https://learn.microsoft.com/en-us/javascript/api/manifest/functionfile?view=word-js-preview)
 * - The FunctionFile HTML can load this JS file, and the function name must match <FunctionName> in the manifest. [1](https://learn.microsoft.com/en-us/javascript/api/manifest/functionfile?view=word-js-preview)
 */

/** Configuration */
const PLACEHOLDER = "{{CODE6}}";             // If present in subject, replace it.
const APPEND_TEXT = " Verification code: ";  // If not present, append this + code.
const MAX_SMS_LEN = 255;                     // ClickSend guidance (warn if exceeded).

/** Generate a 6-digit numeric code string. */
function generate6DigitCode() {
  try {
    if (window.crypto && window.crypto.getRandomValues) {
      const a = new Uint32Array(1);
      window.crypto.getRandomValues(a);
      const code = (a[0] % 900000) + 100000; // 100000-999999
      return String(code);
    }
  } catch (e) {
    // fall through to Math.random
  }
  return String(Math.floor(Math.random() * 900000) + 100000);
}

/** Show a notification in compose if supported, else fallback to console/alert. */
function notify(message, type) {
  // type: "informationalMessage" | "errorMessage"
  try {
    const item = Office.context.mailbox && Office.context.mailbox.item;
    if (item && item.notificationMessages && item.notificationMessages.replaceAsync) {
      item.notificationMessages.replaceAsync("smsVerifyCodeNotice", {
        type: type || "informationalMessage",
        message
      });
      return;
    }
  } catch (e) {
    // ignore
  }

  // Fallbacks
  try { console.log(message); } catch (e) {}
  try { alert(message); } catch (e) {}
}

/**
 * Command function referenced by the manifest's <FunctionName>.
 * Must accept a single `event` parameter and call event.completed() when finished. [1](https://learn.microsoft.com/en-us/javascript/api/manifest/functionfile?view=word-js-preview)
 */
function insertVerifyCode(event) {
  const item = Office.context.mailbox && Office.context.mailbox.item;

  // Guard: this command is intended for compose forms
  if (!item || !item.subject || !item.subject.getAsync) {
    notify("This command is available only while composing a message.", "errorMessage");
    if (event && typeof event.completed === "function") event.completed();
    return;
  }

  const code = generate6DigitCode();

  item.subject.getAsync(function (getResult) {
    if (getResult.status !== Office.AsyncResultStatus.Succeeded) {
      notify("Unable to read the subject. Please try again.", "errorMessage");
      if (event && typeof event.completed === "function") event.completed();
      return;
    }

    const currentSubject = getResult.value || "";
    let newSubject;

    // Replace placeholder if present, else append.
    if (currentSubject.indexOf(PLACEHOLDER) >= 0) {
      newSubject = currentSubject.split(PLACEHOLDER).join(code);
    } else {
      newSubject = (currentSubject.trim() + APPEND_TEXT + code).trim();
    }

    // Warn if subject is too long (ClickSend may split SMS).
    if (newSubject.length > MAX_SMS_LEN) {
      notify(
        `Warning: Subject is ${newSubject.length} characters. ClickSend may split SMS over ${MAX_SMS_LEN} characters.`,
        "informationalMessage"
      );
    }

    item.subject.setAsync(newSubject, function (setResult) {
      if (setResult.status !== Office.AsyncResultStatus.Succeeded) {
        notify("Unable to update the subject. Please try again.", "errorMessage");
      } else {
        notify(`Verification code inserted: ${code}`, "informationalMessage");
      }

      if (event && typeof event.completed === "function") event.completed();
    });
  });
}

/**
 * Initialize and register the command.
 * The FunctionFile HTML must initialize Office.js and define/register the named function used by <FunctionName>. [1](https://learn.microsoft.com/en-us/javascript/api/manifest/functionfile?view=word-js-preview)
 */
Office.onReady(function () {
  // Recommended: explicitly associate the function name from the manifest with this handler.
  if (Office.actions && typeof Office.actions.associate === "function") {
    Office.actions.associate("insertVerifyCode", insertVerifyCode);
  }

  // Optional fallback: expose globally (can help in older clients / debugging).
  window.insertVerifyCode = insertVerifyCode;
});
