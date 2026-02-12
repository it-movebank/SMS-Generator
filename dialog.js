Office.onReady(() => {
  const mobileEl = document.getElementById("mobile");
  const errEl = document.getElementById("err");
  const okBtn = document.getElementById("ok");

  function digitsOnly(raw) {
    return (raw || "").replace(/\D/g, "");
  }

  okBtn.addEventListener("click", () => {
    errEl.textContent = "";
    const mobile = digitsOnly(mobileEl.value);

    if (mobile.length < 9) {
      errEl.textContent = "Please enter a valid mobile number (digits only).";
      return;
    }

    // Send value back to host page (commands.js)
    Office.context.ui.messageParent(JSON.stringify({ mobile }));
  });
});
