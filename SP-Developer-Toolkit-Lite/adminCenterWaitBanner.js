/* global sessionStorage, document, window */
/**
 * Full-screen wait overlay for SharePoint tenant admin pages opened from the
 * toolkit (new tab). sessionStorage is primed by the background service worker;
 * this script also runs at document_start so the overlay survives reloads.
 */
(function () {
  var KEY = "__SPO_TOOLKIT_ADMIN_WAIT__";
  var ROOT_ID = "sp-toolkit-admin-wait-root";
  var settleMs = 3200;
  var failsafeMs = 45000;

  function teardown() {
    try {
      sessionStorage.removeItem(KEY);
    } catch (e) {}
    var n = document.getElementById(ROOT_ID);
    if (n) n.remove();
  }

  function scheduleTeardownOnce() {
    var done = false;
    function finish() {
      if (done) return;
      done = true;
      teardown();
    }
    window.setTimeout(finish, failsafeMs);
    if (document.readyState === "complete") {
      window.setTimeout(finish, settleMs);
    } else {
      window.addEventListener("load", function onLoad() {
        window.removeEventListener("load", onLoad);
        window.setTimeout(finish, settleMs);
      });
    }
  }

  function mount() {
    if (document.getElementById(ROOT_ID)) return;
    var root = document.createElement("div");
    root.id = ROOT_ID;
    root.setAttribute("role", "alertdialog");
    root.setAttribute("aria-modal", "true");
    root.setAttribute("aria-labelledby", "sp-toolkit-admin-wait-title");
    root.setAttribute("aria-describedby", "sp-toolkit-admin-wait-desc");
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
    title.id = "sp-toolkit-admin-wait-title";
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
    msg.id = "sp-toolkit-admin-wait-desc";
    msg.textContent =
      "Loading the SharePoint admin center. Avoid navigating away until the page has finished opening.";
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
    scheduleTeardownOnce();
  }

  function tryPrime() {
    try {
      if (sessionStorage.getItem(KEY) === "1") mount();
    } catch (e) {}
  }

  tryPrime();
})();
