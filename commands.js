/* commands.js */
const PLACEHOLDER = "{{CODE6}}";
const APPEND_TEXT = " Verification code: ";
const MAX_SMS_LEN = 160;
const SMS_DOMAIN = "sms.clicksend.com";
const DIALOG_URL = "https://it-movebank.github.io/SMS-Generator/dialog.html";

function generate6DigitCode() {
  const a = new Uint32Array(1);
  window.crypto.getRandomValues(a);
  return String((a[0] % 900000) + 100000);
}

function notify(message, type) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("smsNotice", {
    type: type || "informationalMessage",
    message: message
  });
}

function insertVerifyCode(event) {
  const item = Office.context.mailbox.item;

  // 1. Check if recipient is already set
  item.to.getAsync((result) => {
    const recipients = result.value || [];
    const hasSms = recipients.some(r => r.emailAddress.endsWith(SMS_DOMAIN));

    if (!hasSms) {
      // 2. No SMS recipient? Open dialog
      Office.context.ui.displayDialogAsync(DIALOG_URL, { height: 40, width: 30 }, (asyncResult) => {
        const dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          const mobile = JSON.parse(arg.message).mobile;
          item.to.addAsync([{ displayName: mobile, emailAddress: `${mobile}@${SMS_DOMAIN}` }]);
          dialog.close();
          // After adding, run insertion
          finishInsertion(event);
        });
      });
    } else {
      finishInsertion(event);
    }
  });
}

function finishInsertion(event) {
  const item = Office.context.mailbox.item;
  const code = generate6DigitCode();

  item.subject.getAsync((result) => {
    let subject = result.value || "";
    let newSubject = subject.includes(PLACEHOLDER) 
      ? subject.replace(PLACEHOLDER, code) 
      : (subject.trim() + APPEND_TEXT + code).trim();

    item.subject.setAsync(newSubject, () => {
      notify(`Code ${code} inserted.`, "informationalMessage");
      event.completed();
    });
  });
}

Office.onReady(() => {
  Office.actions.associate("insertVerifyCode", insertVerifyCode);
});
