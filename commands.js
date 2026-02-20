/* commands.js */
const PLACEHOLDER = "{{CODE6}}";
const APPEND_TEXT = " Verification code: ";
const MAX_SMS_LEN = 160;
const SMS_DOMAIN = "sms.clicksend.com";
const DIALOG_URL = "https://it-movebank.github.io/SMS-Generator/dialog.html";

function safeComplete(event) {
  try {
    if (event && typeof event.completed === "function") {
      event.completed();
    }
  } catch (e) {
    // never throw from completion
  }
}

function generate6DigitCode() {
  try {
    // Prefer crypto if available
    if (window.crypto && window.crypto.getRandomValues) {
      const a = new Uint32Array(1);
      window.crypto.getRandomValues(a);
      return String((a[0] % 900000) + 100000);
    }
  } catch (e) {
    // ignore and fall back
  }
  // Fallback
  return String(Math.floor(Math.random() * 900000) + 100000);
}

function notify(message, type) {
  try {
    const item = Office.context.mailbox && Office.context.mailbox.item;
    if (item && item.notificationMessages && item.notificationMessages.replaceAsync) {
      item.notificationMessages.replaceAsync("smsNotice", {
        type: type || "informationalMessage",
        message: message
      });
      return;
    }
  } catch (e) {
    // ignore
  }

  // Fallbacks if notificationMessages isn't available
  try { console.log(message); } catch (e) {}
  try { alert(message); } catch (e) {}
}

function insertVerifyCode(event) {
  const item = Office.context.mailbox && Office.context.mailbox.item;

  try {
    if (!item || !item.to || !item.to.getAsync) {
      notify("This command is available only while composing a message.", "errorMessage");
      return;
    }

    // 1. Check if recipient is already set
    item.to.getAsync((result) => {
      try {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          notify("Unable to read current recipients. Please try again.", "errorMessage");
          return;
        }

        const recipients = result.value || [];
        const hasSms = recipients.some(r => {
          const addr = (r && r.emailAddress) ? r.emailAddress.toLowerCase() : "";
          return addr.endsWith("@" + SMS_DOMAIN) || addr.endsWith(SMS_DOMAIN);
        });

        if (!hasSms) {
          // 2. No SMS recipient? Open dialog
          Office.context.ui.displayDialogAsync(
            DIALOG_URL,
            { height: 40, width: 30, displayInIframe: true },
            (asyncResult) => {
              if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                notify("Unable to open the mobile number dialog.", "errorMessage");
                safeComplete(event); // ✅ ensure we always complete
                return;
              }

              const dialog = asyncResult.value;

              const closeDialogSafe = () => {
                try { dialog && dialog.close && dialog.close(); } catch (e) {}
              };

              // If the user closes/cancels the dialog, complete the event so Outlook doesn't hang
              dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
                notify("Recipient entry was cancelled.", "informationalMessage");
                closeDialogSafe();
                safeComplete(event); // ✅ critical
              });

              dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                try {
                  let mobile = "";
                  try {
                    mobile = (JSON.parse(arg.message) || {}).mobile || "";
                  } catch (e) {
                    mobile = "";
                  }

                  // Basic cleanup (digits only)
                  mobile = String(mobile).replace(/\D/g, "");

                  if (!mobile || mobile.length < 9) {
                    notify("Please enter a valid mobile number.", "errorMessage");
                    closeDialogSafe();
                    safeComplete(event); // ✅ avoid dangling event
                    return;
                  }

                  const recipient = { displayName: mobile, emailAddress: `${mobile}@${SMS_DOMAIN}` };

                  // Add recipient and only continue once we know the result
                  if (!item.to || !item.to.addAsync) {
                    notify("Unable to add recipient in this Outlook client.", "errorMessage");
                    closeDialogSafe();
                    safeComplete(event);
                    return;
                  }

                  item.to.addAsync([recipient], (addRes) => {
                    if (addRes.status !== Office.AsyncResultStatus.Succeeded) {
                      notify("Failed to add the SMS recipient. Please try again.", "errorMessage");
                      closeDialogSafe();
                      safeComplete(event);
                      return;
                    }

                    closeDialogSafe();
                    // After adding, run insertion
                    finishInsertion(event);
                  });

                } catch (e) {
                  notify("Unexpected error processing the mobile number.", "errorMessage");
                  closeDialogSafe();
                  safeComplete(event); // ✅
                }
              });
            }
          );
        } else {
          finishInsertion(event);
        }
      } catch (e) {
        notify("Unexpected error checking recipients.", "errorMessage");
        safeComplete(event); // ✅
      }
    });

  } catch (e) {
    notify("Unexpected error running the command.", "errorMessage");
    safeComplete(event); // ✅
  }
}

function finishInsertion(event) {
  const item = Office.context.mailbox && Office.context.mailbox.item;

  try {
    if (!item || !item.subject || !item.subject.getAsync) {
      notify("Unable to access the email subject in this client.", "errorMessage");
      safeComplete(event);
      return;
    }

    const code = generate6DigitCode();

    item.subject.getAsync((result) => {
      try {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          notify("Unable to read the subject. Please try again.", "errorMessage");
          safeComplete(event);
          return;
        }

        const subject = result.value || "";
        let newSubject = subject.includes(PLACEHOLDER)
          ? subject.replace(PLACEHOLDER, code)
          : (subject.trim() + APPEND_TEXT + code).trim();

        // Optional warning if subject too long (since MAX_SMS_LEN=160)
        if (newSubject.length > MAX_SMS_LEN) {
          notify(
            `Warning: Subject is ${newSubject.length} characters (limit ${MAX_SMS_LEN}). It may be split.`,
            "informationalMessage"
          );
        }

        item.subject.setAsync(newSubject, (setRes) => {
          if (setRes && setRes.status && setRes.status !== Office.AsyncResultStatus.Succeeded) {
            notify("Unable to update the subject. Please try again.", "errorMessage");
            safeComplete(event);
            return;
          }

          notify(`Code ${code} inserted.`, "informationalMessage");
          safeComplete(event); // ✅ always complete here
        });

      } catch (e) {
        notify("Unexpected error inserting the code.", "errorMessage");
        safeComplete(event); // ✅
      }
    });

  } catch (e) {
    notify("Unexpected error inserting verification code.", "errorMessage");
    safeComplete(event); // ✅
  }
}

Office.onReady(() => {
  Office.actions.associate("insertVerifyCode", insertVerifyCode);
});
