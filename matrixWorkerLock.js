// Full-tab lock overlay for permissions matrix worker tab (progress + mini-game).
(function () {
  const OVERLAY_ID = "spot-matrix-worker-lock";
  const BRAND = "#37ae1c";
  const BRAND_LIGHT = "#86efac";

  let overlayEl = null;
  let barEl = null;
  let pctEl = null;
  let msgEl = null;
  let logEl = null;
  let statusEl = null;
  let gameLoopId = null;
  let snakeRafId = null;
  let snake = null;
  let snakeUserHighScore = 0;
  let snakeAutoHighScore = 0;
  let snakeCurrentScore = 0;
  let activeWorkerReport = "exportCSV";
  let activeWorkerSiteName = "";
  let cachedUserEmail = "";
  let snakeWindowKeyHandler = null;
  let savedDocumentTitle = null;
  let titleGuardTimer = null;
  let themeStorageWired = false;
  let activeWorkerTabTitle = "";
  let sillyModeActive = false;
  let savedFaviconNodes = null;
  const SILLY_TAB_TITLE = "NOTHING TO SEE HERE";
  const SILLY_MODE_KEY = "exportWorkerSillyMode";
  const SNAKE_FAVICON_ID = "spot-mwl-snake-favicon";
  // Compact green snake favicon (data URL — no extra asset file).
  const SNAKE_FAVICON_HREF =
    "data:image/svg+xml," +
    encodeURIComponent(
      "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'>" +
        "<rect width='64' height='64' rx='14' fill='#0f172a'/>" +
        "<path d='M14 34c0-8 6-14 14-14h6c5 0 9 4 9 9s-4 9-9 9H28c-3 0-5 2-5 5s2 5 5 5h18' fill='none' stroke='#37ae1c' stroke-width='7' stroke-linecap='round' stroke-linejoin='round'/>" +
        "<circle cx='44' cy='24' r='3.5' fill='#86efac'/>" +
        "<circle cx='48.5' cy='21.5' r='1.6' fill='#0f172a'/>" +
        "<path d='M50 43c4 2 7 6 8 11' fill='none' stroke='#5fd247' stroke-width='3' stroke-linecap='round'/>" +
      "</svg>"
    );
  const THEME_DARK_CLASS = "spot-mwl-dark";
  const THEME_LIGHT_CLASS = "spot-mwl-light";

  function reportTabLabel(reportKey) {
    const map = {
      permissionsMatrix: "Permissions matrix",
      everythingBagel: "Everything Bagel",
      exportCSV: "List export",
      folderCount: "Folder count",
      pathLengths: "Path lengths"
    };
    return map[reportKey] || "Report export";
  }

  function resolveSiteName(opts) {
    opts = opts || {};
    if (opts.siteName && String(opts.siteName).trim()) return String(opts.siteName).trim();
    try {
      const ctx = window._spPageContextInfo;
      if (ctx) {
        const t = ctx.webTitle || ctx.siteTitle || "";
        if (t && String(t).trim()) return String(t).trim();
      }
    } catch (_) {}
    const fromUrl = function (raw) {
      try {
        const u = new URL(raw, location.origin);
        const segs = u.pathname.replace(/\/$/, "").split("/").filter(Boolean);
        if (!segs.length) return "";
        const sitesIdx = segs.findIndex(function (s) {
          return /^sites$|^teams$/i.test(s);
        });
        if (sitesIdx >= 0 && segs[sitesIdx + 1]) return decodeURIComponent(segs[sitesIdx + 1]);
        return decodeURIComponent(segs[segs.length - 1]);
      } catch (_) {
        return "";
      }
    };
    if (opts.siteUrl) {
      const n = fromUrl(opts.siteUrl);
      if (n) return n;
    }
    const fromLoc = fromUrl(location.href);
    return fromLoc || "SharePoint";
  }

  function buildWorkerTabTitle(opts, silly) {
    if (silly) return SILLY_TAB_TITLE;
    const report = (opts && opts.report) || "exportCSV";
    return reportTabLabel(report) + " | " + resolveSiteName(opts);
  }

  function applySnakeFavicon() {
    try {
      if (savedFaviconNodes == null) {
        savedFaviconNodes = Array.prototype.slice.call(
          document.querySelectorAll("link[rel='icon'], link[rel='shortcut icon']")
        );
        savedFaviconNodes.forEach(function (n) {
          if (n && n.parentNode) n.parentNode.removeChild(n);
        });
      }
      let link = document.getElementById(SNAKE_FAVICON_ID);
      if (!link) {
        link = document.createElement("link");
        link.id = SNAKE_FAVICON_ID;
        link.rel = "icon";
        link.type = "image/svg+xml";
        (document.head || document.documentElement).appendChild(link);
      }
      if (link.getAttribute("href") !== SNAKE_FAVICON_HREF) {
        link.setAttribute("href", SNAKE_FAVICON_HREF);
      }
    } catch (_) {}
  }

  function restoreFavicon() {
    try {
      const snake = document.getElementById(SNAKE_FAVICON_ID);
      if (snake && snake.parentNode) snake.parentNode.removeChild(snake);
      if (Array.isArray(savedFaviconNodes)) {
        savedFaviconNodes.forEach(function (n) {
          try {
            if (n) (document.head || document.documentElement).appendChild(n);
          } catch (_) {}
        });
      }
    } catch (_) {}
    savedFaviconNodes = null;
  }

  function applyWorkerTabChrome(opts, silly) {
    sillyModeActive = !!silly;
    activeWorkerTabTitle = buildWorkerTabTitle(opts, sillyModeActive);
    try {
      if (savedDocumentTitle == null) savedDocumentTitle = document.title;
      if (document.title !== activeWorkerTabTitle) document.title = activeWorkerTabTitle;
    } catch (_) {}
    if (sillyModeActive) applySnakeFavicon();
    else restoreFavicon();
    if (titleGuardTimer) return;
    // SharePoint SPAs often rewrite the title/favicon; keep ours sticky while the lock is up.
    titleGuardTimer = setInterval(function () {
      try {
        if (activeWorkerTabTitle && document.title !== activeWorkerTabTitle) {
          document.title = activeWorkerTabTitle;
        }
        if (sillyModeActive) applySnakeFavicon();
      } catch (_) {}
    }, 800);
  }

  function restoreWorkerTabTitle() {
    if (titleGuardTimer) {
      clearInterval(titleGuardTimer);
      titleGuardTimer = null;
    }
    restoreFavicon();
    sillyModeActive = false;
    activeWorkerTabTitle = "";
    if (savedDocumentTitle != null) {
      try { document.title = savedDocumentTitle; } catch (_) {}
      savedDocumentTitle = null;
    }
  }

  function formatScoreLine(score) {
    return "Score: " + score + " | Auto High: " + snakeAutoHighScore + " | Your High: " + snakeUserHighScore;
  }

  function updateSnakeScoreDisplay(scoreEl, score) {
    snakeCurrentScore = Math.max(0, Math.round(score || 0));
    if (scoreEl) scoreEl.textContent = formatScoreLine(snakeCurrentScore);
  }

  function recordSnakeHighOnFail(failedMode, finalScore) {
    if (failedMode === "auto" && finalScore > snakeAutoHighScore) snakeAutoHighScore = finalScore;
    if (failedMode === "manual" && finalScore > snakeUserHighScore) snakeUserHighScore = finalScore;
  }

  function emailFromLoginName(login) {
    if (!login) return "";
    var s = String(login).trim();
    var pipe = s.lastIndexOf("|");
    var candidate = pipe >= 0 ? s.slice(pipe + 1) : s;
    if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(candidate)) return candidate;
    return "";
  }

  function resolveUserEmailSync() {
    try {
      var ctx = window._spPageContextInfo;
      if (ctx) {
        if (ctx.userEmail && String(ctx.userEmail).indexOf("@") > 0) return String(ctx.userEmail).trim();
        var fromLogin = emailFromLoginName(ctx.userLoginName);
        if (fromLogin) return fromLogin;
      }
    } catch (_) {}
    return "";
  }

  function prefetchUserEmail() {
    var sync = resolveUserEmailSync();
    if (sync) {
      cachedUserEmail = sync;
      return;
    }
    try {
      var ctx = window._spPageContextInfo;
      var webUrl = "";
      if (ctx && ctx.webAbsoluteUrl) webUrl = String(ctx.webAbsoluteUrl).replace(/\/$/, "");
      else webUrl = (location.origin + (location.pathname || "")).replace(/\/$/, "");
      if (!webUrl) return;
      fetch(webUrl + "/_api/web/currentuser?$select=Email,LoginName", {
        credentials: "include",
        headers: { Accept: "application/json;odata=nometadata" }
      }).then(function (resp) {
        if (!resp || !resp.ok) return null;
        return resp.json();
      }).then(function (j) {
        if (!j) return;
        if (j.Email && String(j.Email).indexOf("@") > 0) {
          cachedUserEmail = String(j.Email).trim();
          return;
        }
        var fromLogin = emailFromLoginName(j.LoginName);
        if (fromLogin) cachedUserEmail = fromLogin;
      }).catch(function () {});
    } catch (_) {}
  }

  function openBraggingRightsMailto() {
    var to = cachedUserEmail || resolveUserEmailSync();
    var reportLabel = reportTitle(activeWorkerReport);
    var site = activeWorkerSiteName || resolveSiteName({ siteName: activeWorkerSiteName });
    var subject = "SPOToolkit Auto-Snake bragging rights";
    if (activeWorkerReport === "everythingBagel") {
      subject = "I ran the Everything Bagel and survived Auto-Snake";
    }
    var lines = [
      "Bragging rights unlocked via SPOToolkit.",
      "",
      "Report: " + reportLabel,
      "Site: " + (site || "SharePoint"),
      "",
      "Auto-Snake results:",
      "  Current / final score: " + snakeCurrentScore,
      "  Auto High: " + snakeAutoHighScore,
      "  Your High: " + snakeUserHighScore,
      "",
      activeWorkerReport === "everythingBagel"
        ? "Also: the Everything Bagel finished. The whole schmear. You're welcome."
        : "Export finished. Snake happened. History will remember (this email).",
      "",
      "— sent from the SPOToolkit background export tab"
    ];
    var url =
      "mailto:" + (to ? encodeURIComponent(to) : "") +
      "?subject=" + encodeURIComponent(subject) +
      "&body=" + encodeURIComponent(lines.join("\n"));
    try {
      window.location.href = url;
    } catch (_) {
      try {
        var a = document.createElement("a");
        a.href = url;
        a.rel = "noopener";
        document.documentElement.appendChild(a);
        a.click();
        a.remove();
      } catch (__) {}
    }
  }

  function maybeSendBraggingRights() {
    if (!overlayEl) return;
    var chk = overlayEl.querySelector(".spot-mwl-brag-chk");
    if (!chk || !chk.checked) return;
    if (!cachedUserEmail) cachedUserEmail = resolveUserEmailSync();
    if (cachedUserEmail) {
      openBraggingRightsMailto();
      return;
    }
    // Last chance: wait briefly for currentuser prefetch, then open anyway.
    prefetchUserEmail();
    setTimeout(function () {
      if (!cachedUserEmail) cachedUserEmail = resolveUserEmailSync();
      openBraggingRightsMailto();
    }, 450);
  }

  const CSS =
    "#" + OVERLAY_ID + "{position:fixed;inset:0;z-index:2147483646;font-family:'Segoe UI',system-ui,sans-serif;display:flex;align-items:center;justify-content:center;padding:24px;box-sizing:border-box;overflow:auto;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + "{background:linear-gradient(160deg,#0f172a 0%,#111827 45%,#1a2332 100%);color:#e6edf6;color-scheme:dark;}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + "{background:linear-gradient(160deg,#f5f7fb 0%,#eef2f8 45%,#e8edf5 100%);color:#1c1f4a;color-scheme:light;}" +
    "#" + OVERLAY_ID + " .spot-mwl-card{max-width:520px;width:100%;border-radius:16px;padding:22px 22px 18px;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-card{background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.12);box-shadow:0 24px 64px rgba(0,0,0,.45);}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + " .spot-mwl-card{background:#fff;border:1px solid #e4e6f5;box-shadow:0 18px 48px rgba(28,31,74,.12);}" +
    "#" + OVERLAY_ID + " .spot-mwl-top{display:flex;align-items:flex-start;justify-content:space-between;gap:12px;margin:0 0 8px;}" +
    "#" + OVERLAY_ID + " .spot-mwl-title{margin:0;flex:1;min-width:0;font-size:20px;font-weight:700;letter-spacing:-.02em;}" +
    "#" + OVERLAY_ID + " .spot-mwl-theme{flex:0 0 52px;width:52px;display:flex;align-items:center;padding-top:2px;}" +
    "#" + OVERLAY_ID + " .spot-mwl-theme-btn{width:52px;height:24px;border:none;border-radius:12px;cursor:pointer;position:relative;flex-shrink:0;overflow:hidden;padding:0;background:linear-gradient(90deg,#f5e6c8 0%,#e8d4a8 35%,#3d2a5c 65%,#1a0a2e 100%);box-shadow:inset 0 1px 2px rgba(0,0,0,.2),0 0 0 1px rgba(255,255,255,.15);transition:box-shadow .2s ease;}" +
    "#" + OVERLAY_ID + " .spot-mwl-theme-btn:hover{box-shadow:inset 0 1px 2px rgba(0,0,0,.25),0 0 0 1px rgba(255,255,255,.25);}" +
    "#" + OVERLAY_ID + " .spot-mwl-theme-btn::before{content:'';position:absolute;left:5px;top:50%;transform:translateY(-50%);width:14px;height:14px;background:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23b8860b' stroke-width='1.5' stroke-linecap='round' stroke-linejoin='round'%3E%3Ccircle cx='12' cy='10' r='4'/%3E%3Cpath d='M12 2v2M12 18v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M18 12h2M4.93 19.07l1.41-1.41M17.66 6.34l1.41-1.41'/%3E%3Cpath d='M5 14h14'/%3E%3C/svg%3E\") no-repeat center / contain;opacity:.95;pointer-events:none;}" +
    "#" + OVERLAY_ID + " .spot-mwl-theme-btn::after{content:'';position:absolute;right:5px;top:50%;transform:translateY(-50%);width:14px;height:14px;background:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23c9b8e8' stroke-width='1.5' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z'/%3E%3C/svg%3E\") no-repeat center / contain;opacity:.9;pointer-events:none;}" +
    "#" + OVERLAY_ID + " .spot-mwl-theme-knob{position:absolute;top:2px;left:2px;width:20px;height:20px;border-radius:50%;background:radial-gradient(circle at 30% 30%,#fff 0%,#f0ebe0 50%,#e0d8c8 100%);box-shadow:0 1px 3px rgba(0,0,0,.35),0 0 0 1px rgba(255,255,255,.5);transition:transform .25s cubic-bezier(0.4,0,0.2,1);pointer-events:none;display:flex;align-items:center;justify-content:center;}" +
    "#" + OVERLAY_ID + " .spot-mwl-theme-knob::before{content:'';width:11px;height:11px;background:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23c4952a' stroke-width='1.8' stroke-linecap='round' stroke-linejoin='round'%3E%3Ccircle cx='12' cy='9' r='3'/%3E%3Cpath d='M12 1v1.5M12 19.5V21M4.22 4.22l1.06 1.06M18.72 18.72l1.06 1.06M1 12h1.5M19.5 12H21M4.22 19.78l1.06-1.06M18.72 5.28l1.06-1.06'/%3E%3Cpath d='M4 14h16'/%3E%3C/svg%3E\") no-repeat center / contain;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-theme-knob{transform:translateX(28px);background:radial-gradient(circle at 30% 30%,#c9b8e8 0%,#5c4d7a 50%,#2d2345 100%);box-shadow:0 1px 3px rgba(0,0,0,.5),0 0 0 1px rgba(255,255,255,.15);}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-theme-knob::before{width:10px;height:10px;background:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23e8e0f0' stroke-width='1.5' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z'/%3E%3C/svg%3E\") no-repeat center / contain;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-theme-btn::after{opacity:0;}" +
    "#" + OVERLAY_ID + " .spot-mwl-theme-btn:focus-visible{outline:2px solid " + BRAND + ";outline-offset:2px;}" +
    "#" + OVERLAY_ID + " .spot-mwl-warn{margin:0 0 6px;font-size:13px;font-weight:600;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-warn{color:" + BRAND_LIGHT + ";}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + " .spot-mwl-warn{color:" + BRAND + ";}" +
    "#" + OVERLAY_ID + " .spot-mwl-hint{margin:0 0 10px;font-size:12px;line-height:1.45;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-hint{color:rgba(230,237,246,.72);}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + " .spot-mwl-hint{color:#5a5f8a;}" +
    "#" + OVERLAY_ID + " .spot-mwl-caution{margin:0 0 16px;padding:10px 12px;border-radius:10px;font-size:12px;line-height:1.45;font-weight:600;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-caution{border:1px solid rgba(251,191,36,.35);background:rgba(251,191,36,.1);color:#fde68a;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-caution strong{font-weight:800;color:#fef3c7;}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + " .spot-mwl-caution{border:1px solid #f6d387;background:#fff8e8;color:#92400e;}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + " .spot-mwl-caution strong{font-weight:800;color:#78350f;}" +
    "#" + OVERLAY_ID + " .spot-mwl-bar-wrap{height:10px;border-radius:999px;overflow:hidden;margin-bottom:10px;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-bar-wrap{background:rgba(255,255,255,.08);}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + " .spot-mwl-bar-wrap{background:#e8ecf6;}" +
    "#" + OVERLAY_ID + " .spot-mwl-bar{height:100%;width:0;background:linear-gradient(90deg,#5fd247," + BRAND + ");transition:width .3s ease;border-radius:999px;}" +
    "#" + OVERLAY_ID + " .spot-mwl-head{display:flex;justify-content:space-between;align-items:center;gap:12px;margin-bottom:8px;font-size:13px;font-weight:600;}" +
    "#" + OVERLAY_ID + " .spot-mwl-pct{font-variant-numeric:tabular-nums;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-pct,#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-msg{color:" + BRAND_LIGHT + ";}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + " .spot-mwl-pct,#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + " .spot-mwl-msg{color:" + BRAND + ";}" +
    "#" + OVERLAY_ID + " .spot-mwl-msg{min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}" +
    "#" + OVERLAY_ID + " .spot-mwl-log{max-height:88px;overflow:auto;font:11px/1.4 Consolas,ui-monospace,monospace;white-space:pre-wrap;margin:0 0 14px;padding:8px 10px;border-radius:8px;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-log{color:rgba(230,237,246,.65);background:rgba(0,0,0,.22);}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + " .spot-mwl-log{color:#5a6780;background:#f5f7fc;border:1px solid #e4e6f5;}" +
    "#" + OVERLAY_ID + " .spot-mwl-game{padding-top:14px;margin-top:4px;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-game{border-top:1px solid rgba(255,255,255,.1);}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + " .spot-mwl-game{border-top:1px solid #e4e6f5;}" +
    "#" + OVERLAY_ID + " .spot-mwl-game-head{text-align:center;margin:0 0 10px;}" +
    "#" + OVERLAY_ID + " .spot-mwl-game-title{margin:0 0 4px;font-size:12px;font-weight:500;line-height:1.4;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-game-title{color:rgba(230,237,246,.85);}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + " .spot-mwl-game-title{color:#3d4566;}" +
    "#" + OVERLAY_ID + " .spot-mwl-game-name{font-weight:800;font-size:13px;letter-spacing:.06em;text-transform:uppercase;background:linear-gradient(90deg," + BRAND_LIGHT + "," + BRAND + ");-webkit-background-clip:text;background-clip:text;color:transparent;}" +
    "#" + OVERLAY_ID + " .spot-mwl-game-sub{margin:0;font-size:11px;line-height:1.35;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-game-sub{color:rgba(230,237,246,.58);}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + " .spot-mwl-game-sub{color:#7e8ca3;}" +
    "#" + OVERLAY_ID + " .spot-mwl-crt{display:flex;justify-content:center;}" +
    "#" + OVERLAY_ID + " .spot-mwl-crt-bezel{position:relative;padding:3px;border-radius:8px;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-crt-bezel{background:linear-gradient(180deg,#32363c 0%,#1e2126 55%,#121418 100%);box-shadow:0 4px 14px rgba(0,0,0,.45),inset 0 1px 0 rgba(255,255,255,.08);border:1px solid #0b0d10;}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + " .spot-mwl-crt-bezel{background:linear-gradient(180deg,#f7faf4 0%,#e8f0e0 40%,#d5e3cc 100%);box-shadow:0 8px 22px rgba(55,100,40,.14),inset 0 1px 0 rgba(255,255,255,.9);border:1px solid #b7c9aa;}" +
    "#" + OVERLAY_ID + " .spot-mwl-crt-bezel::before{content:'';position:absolute;top:2px;left:50%;transform:translateX(-50%);width:4px;height:4px;border-radius:50%;opacity:.7;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-crt-bezel::before{background:#101215;box-shadow:inset 0 1px 2px rgba(0,0,0,.75);}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + " .spot-mwl-crt-bezel::before{background:radial-gradient(circle at 35% 35%,#9ae06f,#37ae1c 70%,#1f7a12);box-shadow:0 0 6px rgba(55,174,28,.55);opacity:.95;}" +
    "#" + OVERLAY_ID + " .spot-mwl-crt-screen{position:relative;overflow:hidden;border-radius:4px;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-crt-screen{box-shadow:inset 0 0 18px rgba(0,0,0,.88),0 0 14px rgba(55,174,28,.18);}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + " .spot-mwl-crt-screen{box-shadow:inset 0 0 16px rgba(70,110,50,.12),0 0 0 1px rgba(55,174,28,.18),0 4px 18px rgba(55,174,28,.12);}" +
    "#" + OVERLAY_ID + " .spot-mwl-crt-screen canvas{display:block;border:none;border-radius:0;image-rendering:pixelated;image-rendering:crisp-edges;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-crt-screen canvas{background:#020806;}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + " .spot-mwl-crt-screen canvas{background:#e8f5dc;}" +
    "#" + OVERLAY_ID + " .spot-mwl-crt-scanlines{pointer-events:none;position:absolute;inset:0;z-index:2;animation:spot-mwl-crt-flicker 5s steps(1,end) infinite;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-crt-scanlines{background:repeating-linear-gradient(to bottom,rgba(0,0,0,0) 0,rgba(0,0,0,0) 2px,rgba(0,0,0,.26) 2px,rgba(0,0,0,.26) 4px);mix-blend-mode:multiply;}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + " .spot-mwl-crt-scanlines{background:repeating-linear-gradient(to bottom,rgba(255,255,255,0) 0,rgba(255,255,255,0) 2px,rgba(55,120,40,.06) 2px,rgba(55,120,40,.06) 4px);mix-blend-mode:multiply;opacity:.85;}" +
    "#" + OVERLAY_ID + " .spot-mwl-crt-vignette{pointer-events:none;position:absolute;inset:0;z-index:3;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-crt-vignette{background:radial-gradient(ellipse at center,transparent 52%,rgba(0,0,0,.62) 100%);}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + " .spot-mwl-crt-vignette{background:radial-gradient(ellipse at center,transparent 58%,rgba(80,120,50,.16) 100%);}" +
    "#" + OVERLAY_ID + " .spot-mwl-crt-glare{pointer-events:none;position:absolute;inset:0;z-index:4;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-crt-glare{background:linear-gradient(135deg,rgba(255,255,255,.07) 0%,transparent 32%,transparent 100%);}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + " .spot-mwl-crt-glare{background:linear-gradient(135deg,rgba(255,255,255,.55) 0%,rgba(255,255,255,.12) 28%,transparent 48%);}" +
    "#" + OVERLAY_ID + " .spot-mwl-crt-noise{pointer-events:none;position:absolute;inset:0;z-index:5;background-image:url(\"data:image/svg+xml,%3Csvg viewBox='0 0 256 256' xmlns='http://www.w3.org/2000/svg'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='.95' numOctaves='3' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23n)'/%3E%3C/svg%3E\");animation:spot-mwl-crt-flicker 5s steps(1,end) infinite;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-crt-noise{opacity:.05;mix-blend-mode:soft-light;}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + " .spot-mwl-crt-noise{opacity:.035;mix-blend-mode:multiply;}" +
    "@keyframes spot-mwl-crt-flicker{0%,100%{opacity:.88;}48%{opacity:.94;}49%{opacity:.78;}50%{opacity:.92;}92%{opacity:.9;}93%{opacity:.82;}}" +
    "#" + OVERLAY_ID + " .spot-mwl-score{margin:10px 0 0;font:11px/1.4 Consolas,ui-monospace,monospace;text-align:center;letter-spacing:.02em;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-score{color:#5fd247;text-shadow:0 0 10px rgba(95,210,71,.45),0 0 2px rgba(134,239,172,.35);}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + " .spot-mwl-score{color:#1f7a12;text-shadow:0 1px 0 rgba(255,255,255,.8);}" +
    "#" + OVERLAY_ID + " .spot-mwl-brag{display:flex;align-items:center;justify-content:center;gap:7px;margin:10px 0 0;padding:0;font-size:11px;line-height:1.3;cursor:pointer;user-select:none;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-brag{color:rgba(230,237,246,.42);}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + " .spot-mwl-brag{color:#9aa3b8;}" +
    "#" + OVERLAY_ID + " .spot-mwl-brag:hover{color:inherit;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-brag:hover{color:rgba(230,237,246,.72);}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + " .spot-mwl-brag:hover{color:#5a5f8a;}" +
    "#" + OVERLAY_ID + " .spot-mwl-brag-chk{margin:0;accent-color:" + BRAND + ";width:13px;height:13px;cursor:pointer;opacity:.75;}" +
    "#" + OVERLAY_ID + " .spot-mwl-brag:hover .spot-mwl-brag-chk{opacity:1;}" +
    "#" + OVERLAY_ID + " .spot-mwl-actions{display:flex;gap:10px;margin-top:14px;justify-content:flex-end;}" +
    "#" + OVERLAY_ID + " .spot-mwl-cancel{padding:6px 12px;border-radius:8px;font-size:12px;cursor:pointer;font-family:inherit;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-cancel{border:1px solid rgba(255,255,255,.2);background:rgba(255,255,255,.06);color:#e6edf6;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-cancel:hover{background:rgba(255,255,255,.12);}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + " .spot-mwl-cancel{border:1px solid #cdd0ee;background:#f5f6ff;color:#1c1f4a;}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + " .spot-mwl-cancel:hover{background:#eef0fe;border-color:#b8bddf;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + ".spot-mwl-done .spot-mwl-warn{color:#86efac;}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + ".spot-mwl-done .spot-mwl-warn{color:" + BRAND + ";}" +
    "#" + OVERLAY_ID + ".spot-mwl-failed .spot-mwl-warn{color:#f87171;}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + ".spot-mwl-failed .spot-mwl-warn{color:#dc2626;}" +
    "#" + OVERLAY_ID + " .spot-mwl-close-note{margin:10px 0 0;font-size:11px;text-align:center;}" +
    "#" + OVERLAY_ID + "." + THEME_DARK_CLASS + " .spot-mwl-close-note{color:rgba(230,237,246,.55);}" +
    "#" + OVERLAY_ID + "." + THEME_LIGHT_CLASS + " .spot-mwl-close-note{color:#7e8ca3;}";

  function injectStyles() {
    if (document.getElementById("spot-mwl-styles")) return;
    const s = document.createElement("style");
    s.id = "spot-mwl-styles";
    s.textContent = CSS;
    (document.head || document.documentElement).appendChild(s);
  }

  function updateThemeToggleLabel(isDark) {
    if (!overlayEl) return;
    const btn = overlayEl.querySelector(".spot-mwl-theme-btn");
    if (!btn) return;
    btn.title = isDark ? "Night Owl" : "Early Riser";
    btn.setAttribute(
      "aria-label",
      isDark
        ? "Theme: Night Owl. Click to switch to Early Riser."
        : "Theme: Early Riser. Click to switch to Night Owl."
    );
  }

  function applyTheme(isDark) {
    if (!overlayEl) return;
    overlayEl.classList.remove(THEME_DARK_CLASS, THEME_LIGHT_CLASS);
    overlayEl.classList.add(isDark ? THEME_DARK_CLASS : THEME_LIGHT_CLASS);
    updateThemeToggleLabel(!!isDark);
  }

  function loadAndApplyTheme() {
    try {
      chrome.storage.local.get("darkMode", function (r) {
        applyTheme(!!(r && r.darkMode));
      });
    } catch (_) {
      applyTheme(false);
    }
  }

  function wireThemeStorageListener() {
    if (themeStorageWired) return;
    themeStorageWired = true;
    try {
      chrome.storage.onChanged.addListener(function (changes, areaName) {
        if (areaName !== "local" || !changes.darkMode || !overlayEl) return;
        applyTheme(!!changes.darkMode.newValue);
      });
    } catch (_) {}
  }

  function stopSnake() {
    if (gameLoopId) {
      clearInterval(gameLoopId);
      gameLoopId = null;
    }
    if (snakeRafId) {
      cancelAnimationFrame(snakeRafId);
      snakeRafId = null;
    }
    if (snakeWindowKeyHandler) {
      window.removeEventListener("keydown", snakeWindowKeyHandler, true);
      snakeWindowKeyHandler = null;
    }
    snake = null;
  }

  function startSnake(canvas) {
    stopSnake();
    const ctx = canvas.getContext("2d");
    const cols = 16;
    const rows = 16;
    const TICK_MS = 140;
    const AUTO_IDLE_MS = Math.round(0.8 * cols * TICK_MS);
    const AUTO_LABEL_PULSE_MS = 2200;
    const cell = Math.floor(Math.min(canvas.width / cols, canvas.height / rows));
    const ox = Math.floor((canvas.width - cell * cols) / 2);
    const oy = Math.floor((canvas.height - cell * rows) / 2);
    let dir = { x: 1, y: 0 };
    let nextDir = { x: 1, y: 0 };
    let body = [{ x: 4, y: 8 }, { x: 3, y: 8 }, { x: 2, y: 8 }];
    let food = { x: 10, y: 8 };
    let score = 0;
    let mode = "auto";
    let lastUserInputAt = 0;
    const scoreEl = overlayEl && overlayEl.querySelector(".spot-mwl-score");

    function randFood() {
      for (;;) {
        const p = { x: Math.floor(Math.random() * cols), y: Math.floor(Math.random() * rows) };
        if (!body.some(function (b) { return b.x === p.x && b.y === p.y; })) return p;
      }
    }

    function willCollide(nx, ny, eating) {
      if (nx < 0 || ny < 0 || nx >= cols || ny >= rows) return true;
      var limit = eating ? body.length : body.length - 1;
      for (var i = 0; i < limit; i++) {
        if (body[i].x === nx && body[i].y === ny) return true;
      }
      return false;
    }

    function pickAutoDir() {
      var head = body[0];
      var moves = [
        { x: 0, y: -1 },
        { x: 0, y: 1 },
        { x: -1, y: 0 },
        { x: 1, y: 0 }
      ];
      var valid = [];
      for (var i = 0; i < moves.length; i++) {
        var m = moves[i];
        if (m.x === -dir.x && m.y === -dir.y) continue;
        var nx = head.x + m.x;
        var ny = head.y + m.y;
        var eating = nx === food.x && ny === food.y;
        if (!willCollide(nx, ny, eating)) valid.push(m);
      }
      if (!valid.length) return dir;
      valid.sort(function (a, b) {
        var da = Math.abs(head.x + a.x - food.x) + Math.abs(head.y + a.y - food.y);
        var db = Math.abs(head.x + b.x - food.x) + Math.abs(head.y + b.y - food.y);
        return da - db;
      });
      return valid[0];
    }

    function resetSnake() {
      body = [{ x: 4, y: 8 }, { x: 3, y: 8 }, { x: 2, y: 8 }];
      dir = nextDir = { x: 1, y: 0 };
      score = 0;
      food = randFood();
      updateSnakeScoreDisplay(scoreEl, score);
    }

    function isLightTheme() {
      return !!(overlayEl && overlayEl.classList.contains(THEME_LIGHT_CLASS));
    }

    function draw() {
      var light = isLightTheme();
      var bg = light ? "#e7f6d8" : "#020806";
      var grid = light ? "rgba(40,110,30,.14)" : "rgba(55,174,28,.09)";
      var phosphor = light
        ? (mode === "auto" ? "#1f8a28" : "#16701f")
        : (mode === "auto" ? "#7dff9a" : "#5dff67");
      var hue = light ? 128 : 118;

      ctx.fillStyle = bg;
      ctx.fillRect(0, 0, canvas.width, canvas.height);

      if (light) {
        // Soft LCD wash
        var wash = ctx.createLinearGradient(0, 0, canvas.width, canvas.height);
        wash.addColorStop(0, "rgba(255,255,255,.35)");
        wash.addColorStop(0.45, "rgba(255,255,255,0)");
        wash.addColorStop(1, "rgba(55,140,40,.08)");
        ctx.fillStyle = wash;
        ctx.fillRect(0, 0, canvas.width, canvas.height);
      }

      ctx.strokeStyle = grid;
      ctx.lineWidth = 1;
      for (var gx = 0; gx <= cols; gx++) {
        ctx.beginPath();
        ctx.moveTo(ox + gx * cell + 0.5, oy + 0.5);
        ctx.lineTo(ox + gx * cell + 0.5, oy + rows * cell + 0.5);
        ctx.stroke();
      }
      for (var gy = 0; gy <= rows; gy++) {
        ctx.beginPath();
        ctx.moveTo(ox + 0.5, oy + gy * cell + 0.5);
        ctx.lineTo(ox + cols * cell + 0.5, oy + gy * cell + 0.5);
        ctx.stroke();
      }

      var bodyLen = body.length;
      for (var si = bodyLen - 1; si >= 0; si--) {
        var seg = body[si];
        var t = bodyLen > 1 ? si / (bodyLen - 1) : 0;
        var fade = 1 - t;
        var fadeCurve = fade * fade * (3 - 2 * fade);
        var alpha = light ? (0.35 + fadeCurve * 0.65) : (0.1 + fadeCurve * 0.9);
        var segSat = light
          ? Math.round(58 + fadeCurve * 28)
          : Math.round(48 + fadeCurve * 52);
        var segLight = light
          ? Math.round(28 + fadeCurve * 22)
          : Math.round(14 + fadeCurve * 50);
        var segBlur = light ? (0.5 + fadeCurve * 3) : (1 + fadeCurve * 11);
        var inset = 1 + Math.round(t * 1.2);
        var segSize = Math.max(2, cell - 2 - Math.round(t * 1.4));
        var sx = ox + seg.x * cell + inset;
        var sy = oy + seg.y * cell + inset;

        ctx.shadowColor = phosphor;
        ctx.globalAlpha = alpha * (light ? 0.25 : 0.4);
        ctx.shadowBlur = segBlur + fadeCurve * (light ? 2 : 8);
        ctx.fillStyle = "hsl(" + hue + ", " + Math.round(segSat * 0.75) + "%, " + Math.round(segLight * (light ? 0.85 : 0.55)) + "%)";
        ctx.fillRect(sx, sy, segSize, segSize);

        ctx.globalAlpha = alpha * (light ? (0.7 + fadeCurve * 0.3) : (0.55 + fadeCurve * 0.45));
        ctx.shadowBlur = segBlur;
        ctx.fillStyle = "hsl(" + hue + ", " + segSat + "%, " + segLight + "%)";
        ctx.fillRect(sx, sy, segSize, segSize);

        if (light && si === 0) {
          // Tiny highlight on the head
          ctx.globalAlpha = 0.55;
          ctx.shadowBlur = 0;
          ctx.fillStyle = "rgba(255,255,255,.45)";
          ctx.fillRect(sx + 1, sy + 1, Math.max(1, Math.floor(segSize * 0.35)), Math.max(1, Math.floor(segSize * 0.35)));
        }
      }
      ctx.globalAlpha = 1;
      ctx.shadowBlur = 0;

      var foodX = ox + food.x * cell + 2;
      var foodY = oy + food.y * cell + 2;
      var foodSize = cell - 4;
      if (light) {
        ctx.shadowColor = "rgba(220,90,40,.45)";
        ctx.shadowBlur = 8;
        ctx.fillStyle = "#ef6c2f";
        ctx.fillRect(foodX, foodY, foodSize, foodSize);
        ctx.shadowBlur = 0;
        ctx.fillStyle = "rgba(255,220,160,.7)";
        ctx.fillRect(foodX + 1, foodY + 1, Math.max(1, foodSize - 4), Math.max(1, Math.floor(foodSize * 0.35)));
      } else {
        ctx.shadowBlur = 14;
        ctx.fillStyle = BRAND_LIGHT;
        ctx.fillRect(foodX, foodY, foodSize, foodSize);
        ctx.shadowBlur = 0;
      }

      if (mode === "auto") {
        var pulse = 0.5 + 0.5 * Math.sin((Date.now() * Math.PI * 2) / AUTO_LABEL_PULSE_MS);
        ctx.font = "10px Consolas,ui-monospace,monospace";
        if (light) {
          var lightAlpha = 0.45 + pulse * 0.5;
          ctx.globalAlpha = lightAlpha;
          ctx.shadowColor = "rgba(31,122,18,.35)";
          ctx.shadowBlur = 4 + pulse * 6;
          ctx.fillStyle = "hsl(128, 62%, " + Math.round(28 + pulse * 10) + "%)";
          ctx.fillText("AUTO", ox + 4, oy + 12);
          ctx.shadowBlur = 0;
          ctx.globalAlpha = 1;
        } else {
          var autoAlpha = 0.38 + pulse * 0.62;
          var letterSat = Math.round(52 + pulse * 48);
          var letterLight = Math.round(26 + pulse * 44);
          var glowSat = Math.round(70 + pulse * 30);
          var glowLight = Math.round(42 + pulse * 38);
          ctx.shadowColor = "hsl(118, " + glowSat + "%, " + glowLight + "%)";
          ctx.globalAlpha = autoAlpha * (0.4 + pulse * 0.35);
          ctx.shadowBlur = 16 + pulse * 24;
          ctx.fillStyle = "hsl(118, " + Math.round(letterSat * 0.85) + "%, " + Math.round(letterLight * 0.75) + "%)";
          ctx.fillText("AUTO", ox + 4, oy + 12);
          ctx.globalAlpha = autoAlpha;
          ctx.shadowBlur = 8 + pulse * 16;
          ctx.fillStyle = "hsl(118, " + letterSat + "%, " + letterLight + "%)";
          ctx.fillText("AUTO", ox + 4, oy + 12);
          ctx.shadowBlur = 0;
          ctx.globalAlpha = 1;
        }
      }
    }

    function renderLoop() {
      if (!snake) return;
      draw();
      snakeRafId = requestAnimationFrame(renderLoop);
    }

    function tick() {
      if (mode === "manual" && lastUserInputAt && Date.now() - lastUserInputAt >= AUTO_IDLE_MS) {
        mode = "auto";
      }
      if (mode === "auto") {
        nextDir = pickAutoDir();
      }
      dir = nextDir;
      var head = { x: body[0].x + dir.x, y: body[0].y + dir.y };
      var eating = head.x === food.x && head.y === food.y;
      if (willCollide(head.x, head.y, eating)) {
        recordSnakeHighOnFail(mode, score);
        resetSnake();
        return;
      }
      body.unshift(head);
      if (eating) {
        score++;
        updateSnakeScoreDisplay(scoreEl, score);
        food = randFood();
      } else {
        body.pop();
      }
    }

    function onKey(e) {
      var k = e.key;
      if (k !== "ArrowUp" && k !== "ArrowDown" && k !== "ArrowLeft" && k !== "ArrowRight") return;
      e.preventDefault();
      mode = "manual";
      lastUserInputAt = Date.now();
      if (k === "ArrowUp" && dir.y === 0) nextDir = { x: 0, y: -1 };
      else if (k === "ArrowDown" && dir.y === 0) nextDir = { x: 0, y: 1 };
      else if (k === "ArrowLeft" && dir.x === 0) nextDir = { x: -1, y: 0 };
      else if (k === "ArrowRight" && dir.x === 0) nextDir = { x: 1, y: 0 };
    }

    snakeWindowKeyHandler = onKey;
    window.addEventListener("keydown", snakeWindowKeyHandler, true);
    snake = { canvas: canvas, onKey: onKey };
    updateSnakeScoreDisplay(scoreEl, score);
    snakeRafId = requestAnimationFrame(renderLoop);
    gameLoopId = setInterval(tick, TICK_MS);
  }

  function reportTitle(reportKey) {
    const map = {
      permissionsMatrix: "Permissions matrix export",
      everythingBagel: "Everything Bagel export",
      exportCSV: "List / library export",
      folderCount: "Folder count report",
      pathLengths: "Path length report"
    };
    return map[reportKey] || "Report export";
  }

  function reportBannerCopy(reportKey) {
    if (reportKey === "everythingBagel") {
      return {
        warn: "You are now running the Everything Bagel report.",
        hint:
          "Legendary status unlocked: permissions matrix, group members, sharing links, all items, folder counts, and path lengths — the whole schmear in one workbook. " +
          "Runtime scales with how deep your site jungle is (lists, libraries, unique perms, nested folders). Small site? Snack break. Sprawling estate? Pack a lunch and a high score. " +
          "Your original SharePoint tab stays free — Auto-Snake is standing by. This tab closes when the bagel is baked."
      };
    }
    return {
      warn: "Background export tab",
      hint:
        "This tab runs your export in the background so your original SharePoint tab stays free. Play Auto-Snake while you wait — this tab closes automatically when finished."
    };
  }

  function show(reportKeyOrOpts) {
    const opts =
      reportKeyOrOpts && typeof reportKeyOrOpts === "object"
        ? reportKeyOrOpts
        : { report: reportKeyOrOpts || "exportCSV" };
    if (!opts.report) opts.report = "exportCSV";
    activeWorkerReport = opts.report;
    activeWorkerSiteName = resolveSiteName(opts);
    snakeCurrentScore = 0;
    snakeUserHighScore = 0;
    snakeAutoHighScore = 0;
    prefetchUserEmail();
    injectStyles();
    wireThemeStorageListener();
    const title = reportTitle(opts.report);
    const banner = reportBannerCopy(opts.report);
    function afterSillyResolved(silly) {
      applyWorkerTabChrome(opts, silly);
    }
    try {
      chrome.storage.local.get(SILLY_MODE_KEY, function (st) {
        afterSillyResolved(!!(st && st[SILLY_MODE_KEY]));
      });
    } catch (_) {
      afterSillyResolved(false);
    }
    if (overlayEl && overlayEl.parentNode) {
      overlayEl.classList.remove("spot-mwl-done", "spot-mwl-failed");
      const titleEl = overlayEl.querySelector(".spot-mwl-title");
      if (titleEl) titleEl.textContent = title;
      const warnEl = overlayEl.querySelector(".spot-mwl-warn");
      if (warnEl) warnEl.textContent = banner.warn;
      const hintEl = overlayEl.querySelector(".spot-mwl-hint");
      if (hintEl) hintEl.textContent = banner.hint;
      const bragChk = overlayEl.querySelector(".spot-mwl-brag-chk");
      if (bragChk) bragChk.checked = false;
      const scoreElReuse = overlayEl.querySelector(".spot-mwl-score");
      if (scoreElReuse) scoreElReuse.textContent = formatScoreLine(0);
      loadAndApplyTheme();
      return;
    }
    overlayEl = document.createElement("div");
    overlayEl.id = OVERLAY_ID;
    overlayEl.className = THEME_DARK_CLASS;
    overlayEl.innerHTML =
      '<div class="spot-mwl-card">' +
      '<div class="spot-mwl-top">' +
      '<h1 class="spot-mwl-title">' + title + '</h1>' +
      '<div class="spot-mwl-theme">' +
      '<button type="button" class="spot-mwl-theme-btn" title="Early Riser" aria-label="Theme: Early Riser. Click to switch to Night Owl.">' +
      '<span class="spot-mwl-theme-knob" aria-hidden="true"></span></button></div></div>' +
      '<p class="spot-mwl-warn">' + banner.warn + '</p>' +
      '<p class="spot-mwl-hint">' + banner.hint + '</p>' +
      '<p class="spot-mwl-caution"><strong>Do not close or refresh this tab</strong> while the export is running. Closing or refreshing stops the job and you will need to start the export again.</p>' +
      '<div class="spot-mwl-bar-wrap"><div class="spot-mwl-bar"></div></div>' +
      '<div class="spot-mwl-head"><span class="spot-mwl-msg">Starting…</span><span class="spot-mwl-pct">0%</span></div>' +
      '<pre class="spot-mwl-log"></pre>' +
      '<div class="spot-mwl-game">' +
      '<div class="spot-mwl-game-head">' +
      '<p class="spot-mwl-game-title">While you wait — <span class="spot-mwl-game-name">Auto-Snake</span></p>' +
      '<p class="spot-mwl-game-sub">(arrow keys to take over)</p></div>' +
      '<div class="spot-mwl-crt"><div class="spot-mwl-crt-bezel"><div class="spot-mwl-crt-screen">' +
      '<canvas width="256" height="256" aria-label="Snake mini-game"></canvas>' +
      '<div class="spot-mwl-crt-scanlines" aria-hidden="true"></div>' +
      '<div class="spot-mwl-crt-vignette" aria-hidden="true"></div>' +
      '<div class="spot-mwl-crt-glare" aria-hidden="true"></div>' +
      '<div class="spot-mwl-crt-noise" aria-hidden="true"></div></div></div></div>' +
      '<p class="spot-mwl-score">Score: 0 | Auto High: 0 | Your High: 0</p>' +
      '<label class="spot-mwl-brag"><input type="checkbox" class="spot-mwl-brag-chk" /> Send me bragging rights</label></div>' +
      '<div class="spot-mwl-actions"><button type="button" class="spot-mwl-cancel">Cancel export</button></div>' +
      '<p class="spot-mwl-close-note"></p></div>';
    document.documentElement.appendChild(overlayEl);
    barEl = overlayEl.querySelector(".spot-mwl-bar");
    pctEl = overlayEl.querySelector(".spot-mwl-pct");
    msgEl = overlayEl.querySelector(".spot-mwl-msg");
    logEl = overlayEl.querySelector(".spot-mwl-log");
    statusEl = overlayEl.querySelector(".spot-mwl-close-note");
    overlayEl.querySelector(".spot-mwl-cancel").addEventListener("click", function () {
      const cancelBtn = overlayEl.querySelector(".spot-mwl-cancel");
      if (cancelBtn) {
        cancelBtn.disabled = true;
        cancelBtn.textContent = "Cancelling…";
      }
      if (msgEl) msgEl.textContent = "Cancelling export…";
      // Signal page-context export immediately (don't wait for background round-trip).
      try {
        window.__SPOToolkitExportCancel = true;
        window.__SPOToolkitExportRunning = false;
        const s = document.createElement("script");
        s.textContent = "window.__SPOToolkitExportCancel=true;try{window.dispatchEvent(new CustomEvent('spotoolkit-export-cancel'));}catch(e){}";
        (document.documentElement || document.head).appendChild(s);
        s.remove();
      } catch (_) {}
      try { window.postMessage({ __spcsv: true, type: "SPCSVExportCancel" }, "*"); } catch (_) {}
      try { chrome.runtime.sendMessage({ type: "SPCSVExportCancel" }); } catch (_) {}
    });
    const themeBtn = overlayEl.querySelector(".spot-mwl-theme-btn");
    if (themeBtn) {
      themeBtn.addEventListener("click", function (e) {
        e.preventDefault();
        e.stopPropagation();
        const nextDark = !overlayEl.classList.contains(THEME_DARK_CLASS);
        applyTheme(nextDark);
        try { chrome.storage.local.set({ darkMode: nextDark }); } catch (_) {}
      });
    }
    loadAndApplyTheme();
    const scoreElInit = overlayEl.querySelector(".spot-mwl-score");
    if (scoreElInit) scoreElInit.textContent = formatScoreLine(0);
    startSnake(overlayEl.querySelector(".spot-mwl-crt-screen canvas"));
  }

  function update(opts) {
    if (!overlayEl) return;
    opts = opts || {};
    if (opts.percent != null && barEl && pctEl) {
      const p = Math.max(0, Math.min(100, Math.round(opts.percent)));
      barEl.style.width = p + "%";
      pctEl.textContent = p + "%";
    }
    if (opts.message && msgEl) msgEl.textContent = String(opts.message);
    if (opts.logLine && logEl) {
      const line = String(opts.logLine).trim();
      if (line) {
        const ts = new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit", second: "2-digit" });
        const prev = logEl.textContent || "";
        logEl.textContent = (prev ? prev + "\n" : "") + "[" + ts + "] " + line;
        logEl.scrollTop = logEl.scrollHeight;
      }
    }
  }

  function finish(success, message, autoCloseMs) {
    if (!overlayEl) return;
    stopSnake();
    const cancelled = /cancel/i.test(String(message || ""));
    overlayEl.classList.toggle("spot-mwl-done", !!success);
    overlayEl.classList.toggle("spot-mwl-failed", !success && !cancelled);
    const warn = overlayEl.querySelector(".spot-mwl-warn");
    if (warn) {
      warn.textContent = success ? "Export complete" : (cancelled ? "Export cancelled" : "Export failed");
    }
    if (msgEl && message) msgEl.textContent = String(message);
    if (barEl && success) barEl.style.width = "100%";
    if (pctEl && success) pctEl.textContent = "100%";
    if (statusEl) {
      const sec = Math.max(1, Math.round((autoCloseMs || 5000) / 1000));
      statusEl.textContent = "Closing this tab in " + sec + " second" + (sec === 1 ? "" : "s") + "…";
    }
    const cancelBtn = overlayEl.querySelector(".spot-mwl-cancel");
    if (cancelBtn) cancelBtn.style.display = "none";
    if (success) {
      // Let the "complete" UI paint first, then hand off to the mail client.
      setTimeout(function () {
        maybeSendBraggingRights();
      }, 250);
    }
  }

  function hide() {
    stopSnake();
    restoreWorkerTabTitle();
    if (overlayEl && overlayEl.parentNode) overlayEl.parentNode.removeChild(overlayEl);
    overlayEl = barEl = pctEl = msgEl = logEl = statusEl = null;
  }

  function isActive() {
    return !!(overlayEl && overlayEl.parentNode);
  }

  window.SPOToolkitMatrixWorkerLock = {
    show: show,
    update: update,
    finish: finish,
    hide: hide,
    isActive: isActive
  };
})();
