/* commands.js
 * Outlook add-in command functions:
 *  - setSmsRecipient: prompts for a mobile number and adds <mobile>@sms.clicksend.com to To
 *  - insertVerifyCode: generates 6-digit code and inserts into Subject (replaces {{CODE6}} or appends)
 */

const PLACEHOLDER = "{{CODE6}}";
const APPEND_TEXT = " Verification code: ";
const MAX_SMS_LEN = 255;

const SMS_DOMAIN = "sms.clicksend.com";
const DIALOG_URL = "https://it-movebank.github.io/SMS-Generator/dialog.html";

function generate6DigitCode() {
  try {
    if (window.crypto && window.crypto.getRandomValues) {
      const a = new Uint32Array(1);
      window.crypto.getRandomValues(a);
      const code = (a[0] % 900000) + 100000;
      return String(code);
    }
  } catch (e) {}
  return String(Math.floor(Math.random() * 900000) + 100000);
}

function digitsOnly(raw) {
  return (raw || "").replace(/\D/g, "");
}

function toSmsEmail(mobileDigits) {
  return `${mobileDigits}@${SMS_DOMAIN}`;
}

function notify(message, type) {
  try {
    const item = Office.context.mailbox && Office.context.mailbox.item;
    if (item && item.notificationMessages && item.notificationMessages.replaceAsync) {
      item.notificationMessages.replaceAsync("smsVerifyCodeNotice", {
        type: type || "informationalMessage",
        message
      });
      return;
    }
  } catch (e) {}
  try { console.log(message); } catch (e) {}
  try { alert(message); } catch (e) {}
}

function ensureSmsRecipient(callback) {
  const item = Office.context.mailbox && Office.context.mailbox.item;
  if (!item || !item.to || !item.to.getAsync) {
    callback(new Error("This action is only available while composing a message."));
    return;
  }

  item.to.getAsync((res) => {
    if (res.status !== Office.AsyncResultStatus.Succeeded) {
      callback(new Error("Unable to read current To recipients."));
      return;
    }

    const existing = (res.value || []).map(r => (r.emailAddress || "").toLowerCase());
    const hasSms = existing.some(a => a.endsWith(`@${SMS_DOMAIN}`));

    if (hasSms) {
      callback(null);
      return;
    }

    Office.context.ui.displayDialogAsync(
      DIALOG_URL,
      { height: 30, width: 30, displayInIframe: true },
      (asyncResult) => {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          callback(new Error(asyncResult.error && asyncResult.error.message ? asyncResult.error.message : "Unable to open mobile input dialog."));
          return;
        }

        const dialog = asyncResult.value;

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          let mobileDigits = "";
          try {
            const payload = JSON.parse(arg.message || "{}");
            mobileDigits = digitsOnly(payload.mobile);
          } catch (e) {}

          if (!mobileDigits || mobileDigits.length < 9) {
            try { dialog.close(); } catch (e) {}
            callback(new Error("Invalid mobile number."));
            return;
          }

          const smsEmail = toSmsEmail(mobileDigits);

          // Add as EmailAddressDetails so displayName can be digits while SMTP is correct.
          item.to.addAsync([{ displayName: mobileDigits, emailAddress: smsEmail }], (addRes) => {
            try { dialog.close(); } catch (e) {}

            if (addRes.status !== Office.AsyncResultStatus.Succeeded) {
              callback(new Error("Failed to add SMS recipient to To field."));
            } else {
              notify(`Added SMS recipient: ${smsEmail}`, "informationalMessage");
              callback(null);
            }
          });
        });
      }
    );
  });
}

function setSmsRecipient(event) {
  ensureSmsRecipient((err) => {
    if (err) notify(err.message, "errorMessage");
    if (event && typeof event.completed === "function") event.completed();
  });
}

function insertVerifyCode(event) {
  ensureSmsRecipient((err) => {
    if (err) {
      notify(err.message, "errorMessage");
      if (event && typeof event.completed === "function") event.completed();
      return;
    }

    const item = Office.context.mailbox && Office.context.mailbox.item;
    if (!item || !item.subject || !item.subject.getAsync) {
      notify("This command is available only while composing a message.", "errorMessage");
      if (event && typeof event.completed === "function") event.completed();
      return;
    }

    const code = generate6DigitCode();

    item.subject.getAsync((getResult) => {
      if (getResult.status !== Office.AsyncResultStatus.Succeeded) {
        notify("Unable to read the subject. Please try again.", "errorMessage");
        if (event && typeof event.completed === "function") event.completed();
        return;
      }

      const currentSubject = getResult.value || "";
      let newSubject;

      if (currentSubject.indexOf(PLACEHOLDER) >= 0) {
        newSubject = currentSubject.split(PLACEHOLDER).join(code);
      } else {
        newSubject = (currentSubject.trim() + APPEND_TEXT + code).trim();
      }

      if (newSubject.length > MAX_SMS_LEN) {
        notify(
          `Warning: Subject is ${newSubject.length} characters. ClickSend may split SMS over ${MAX_SMS_LEN} characters.`,
          "informationalMessage"
        );
      }

      item.subject.setAsync(newSubject, (setResult) => {
        if (setResult.status !== Office.AsyncResultStatus.Succeeded) {
          notify("Unable to update the subject. Please try again.", "errorMessage");
        } else {
          notify(`Verification code inserted: ${code}`, "informationalMessage");
        }

        if (event && typeof event.completed === "function") event.completed();
      });
    });
  });
}

Office.onReady(() => {
  if (Office.actions && typeof Office.actions.associate === "function") {
    Office.actions.associate("insertVerifyCode", insertVerifyCode);
    Office.actions.associate("setSmsRecipient", setSmsRecipient);
  }
  window.insertVerifyCode = insertVerifyCode;
  window.setSmsRecipient = setSmsRecipient;
});
