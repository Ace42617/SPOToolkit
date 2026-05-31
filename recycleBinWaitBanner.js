/* global sessionStorage, document, window */
/**
 * Re-applies the recycle-bin wait overlay on every full load while the
 * background sequence has primed sessionStorage (same tab, same origin).
 * Runs at document_start so it beats first paint after refresh/navigate.
 */
(function () {
  var KEY = "__SPO_TOOLKIT_RB_WAIT__";
  var TS_KEY = "__SPO_TOOLKIT_RB_WAIT_TS__";
  var ROOT_ID = "sp-toolkit-rb-wait-root";
  var maxAgeMs = 45000;

  function teardown() {
    try {
      sessionStorage.removeItem(KEY);
      sessionStorage.removeItem(TS_KEY);
    } catch (e) {}
    var n = document.getElementById(ROOT_ID);
    if (n) n.remove();
  }

  function remainingWaitMs() {
    try {
      if (sessionStorage.getItem(KEY) !== "1") return 0;
      var ts = parseInt(sessionStorage.getItem(TS_KEY) || "", 10);
      if (!ts || isNaN(ts)) return 0;
      var remaining = maxAgeMs - (Date.now() - ts);
      return remaining > 0 ? remaining : 0;
    } catch (e) {
      return 0;
    }
  }

  function mount(remainingMs) {
    if (document.getElementById(ROOT_ID)) return;
    var root = document.createElement("div");
    root.id = ROOT_ID;
    root.setAttribute("role", "alertdialog");
    root.setAttribute("aria-modal", "true");
    root.setAttribute("aria-labelledby", "sp-toolkit-rb-wait-title");
    root.setAttribute("aria-describedby", "sp-toolkit-rb-wait-desc");
    Object.assign(root.style, {
      position: "fixed",
      inset: "0",
      zIndex: "2147483647",
      background: "rgba(12, 14, 22, 0.94)",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      fontFamily: '"Segoe UI", system-ui, -apple-system, sans-serif',
      boxSizing: "border-box",
      backdropFilter: "blur(2px)",
    });

    var card = document.createElement("div");
    Object.assign(card.style, {
      maxWidth: "440px",
      margin: "20px",
      padding: "32px 36px",
      background: "linear-gradient(165deg, #1e2436 0%, #151a28 100%)",
      border: "4px solid #0d9488",
      borderRadius: "14px",
      boxShadow:
        "0 28px 60px rgba(0,0,0,.55), 0 0 0 1px rgba(45,212,191,.22)",
      color: "#f1f5f9",
      textAlign: "center",
    });

    var iconRow = document.createElement("div");
    iconRow.textContent = "⏳";
    Object.assign(iconRow.style, {
      fontSize: "48px",
      lineHeight: "1",
      marginBottom: "16px",
    });

    var title = document.createElement("div");
    title.id = "sp-toolkit-rb-wait-title";
    title.textContent = "Please wait — do not close this tab";
    Object.assign(title.style, {
      fontSize: "clamp(20px, 4vw, 26px)",
      fontWeight: "800",
      marginBottom: "14px",
      color: "#5eead4",
      letterSpacing: "0.02em",
      lineHeight: "1.25",
    });

    var msg = document.createElement("p");
    msg.id = "sp-toolkit-rb-wait-desc";
    msg.textContent =
      "Opening the site collection second-stage recycle bin. Do not click, type, refresh, or navigate away until loading finishes.";
    Object.assign(msg.style, {
      fontSize: "15px",
      lineHeight: "1.5",
      margin: "0",
      color: "rgba(248, 250, 252, 0.92)",
    });

    card.appendChild(iconRow);
    card.appendChild(title);
    card.appendChild(msg);
    root.appendChild(card);
    (document.body || document.documentElement).appendChild(root);
    window.setTimeout(teardown, Math.max(1000, remainingMs || 0));
  }

  function tryPrime() {
    var remaining = remainingWaitMs();
    if (remaining > 0) mount(remaining);
    else teardown();
  }

  tryPrime();
})();

