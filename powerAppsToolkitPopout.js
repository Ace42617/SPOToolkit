/**
 * In-page Power Apps toolkit popout — mirrors SharePoint compass launcher
 * size/position (shared storage) and adds a rocket FAB animation.
 */
(function () {
  "use strict";

  if (window.__spoToolkitPowerAppsPopoutLoaded) return;
  window.__spoToolkitPowerAppsPopoutLoaded = true;

  const ROOT_ID = "spo-toolkit-pa-launcher";
  const PANEL_ID = "spo-toolkit-pa-panel";
  const STYLE_ID = "spo-toolkit-pa-launcher-style";
  const FRAME_ID = "spo-toolkit-pa-frame";
  /** Shared with SharePoint compass launcher so size/location stay in sync. */
  const POSITION_KEY = "listsLauncherPosition";
  const ICON_SCALE_KEY = "listsLauncherIconScale";
  const LEGACY_POSITION_KEY = "powerAppsPopoutPosition";
  const CONFIG_KEY = "levelup_extension_config";
  const THEME_KEY = "levelup-theme-mode";

  const DISPLAY_MODE_PRESETS = {
    default: {
      showRecentlyUsed: true,
      showFavorites: true,
      showActionSections: true,
      showCustomCommands: true,
      showImpersonation: true,
      showGitHubIntegration: true,
      showFormSection: true,
      showNavigationSection: true,
      showDebuggingSection: true,
    },
    simple: {
      showRecentlyUsed: false,
      showFavorites: false,
      showActionSections: true,
      showCustomCommands: false,
      showImpersonation: false,
      showGitHubIntegration: false,
      showFormSection: true,
      showNavigationSection: true,
      showDebuggingSection: false,
    },
  };

  const FAB_SIZE_DEFAULT = 58;
  const DEFAULT_ICON_SCALE = 1.28;
  const SCALE_MIN = 0.75;
  const SCALE_MAX = 1.6;
  const PANEL_GAP = 12;
  const VIEW_MARGIN = 14;
  /** Match SharePoint compass panel open/close fade timings. */
  const PANEL_FADE_IN_S = 0.2;
  const PANEL_FADE_OUT_S = 0.12;
  const PANEL_FADE_OUT_MS = Math.round(PANEL_FADE_OUT_S * 1000);

  const GEAR_SVG =
    '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">' +
    '<path fill="currentColor" d="M11.983 3.5a1 1 0 0 1 .894.553l.535 1.06a6.96 6.96 0 0 1 1.221.508l1.107-.366a1 1 0 0 1 1.09.268l1.061 1.06a1 1 0 0 1 .268 1.09l-.366 1.108c.18.397.34.805.478 1.22l1.09.55a1 1 0 0 1 .552.894v1.5a1 1 0 0 1-.553.894l-1.06.535a6.93 6.93 0 0 1-.508 1.221l.366 1.107a1 1 0 0 1-.268 1.09l-1.06 1.061a1 1 0 0 1-1.09.268l-1.108-.366a6.955 6.955 0 0 1-1.22.478l-.55 1.09a1 1 0 0 1-.894.552h-1.5a1 1 0 0 1-.894-.553l-.535-1.06a6.93 6.93 0 0 1-1.221-.508l-1.107.366a1 1 0 0 1-1.09-.268L5.13 18.74a1 1 0 0 1-.268-1.09l.366-1.108a6.95 6.95 0 0 1-.478-1.22l-1.09-.55a1 1 0 0 1-.552-.894v-1.5a1 1 0 0 1 .553-.894l1.06-.535a6.955 6.955 0 0 1 .508-1.221l-.366-1.107a1 1 0 0 1 .268-1.09L6.19 5.38a1 1 0 0 1 1.09-.268l1.108.366c.397-.18.805-.34 1.22-.478l.55-1.09a1 1 0 0 1 .894-.552h1.5zm.017 5.25a3.5 3.5 0 1 0 0 7.001 3.5 3.5 0 0 0 0-7.001z"/>' +
    "</svg>";

  let expanded = false;
  let rootEl = null;
  let fabPx = Math.round(FAB_SIZE_DEFAULT * DEFAULT_ICON_SCALE);
  let resizeTimer = null;
  let closeAnimTimer = null;

  function isDynamicsRuntime() {
    try {
      const win = window;
      if (win.Xrm?.Utility?.getGlobalContext) {
        const version = win.Xrm.Utility.getGlobalContext().getVersion();
        if (version && String(version).startsWith("9.")) return true;
      }
    } catch (_) {}
    try {
      return Array.from(document.querySelectorAll("script[src]")).some((script) => {
        const src = script.src || "";
        return (
          src.indexOf("/uclient/scripts") !== -1 ||
          src.indexOf("/_static/_common/scripts/PageLoader.js") !== -1 ||
          src.indexOf("/_static/_common/scripts/crminternalutility.js") !== -1
        );
      });
    } catch (_) {
      return false;
    }
  }

  function normalizeScale(value) {
    const parsed = Number(value);
    if (!Number.isFinite(parsed)) return DEFAULT_ICON_SCALE;
    return Math.min(SCALE_MAX, Math.max(SCALE_MIN, parsed));
  }

  function applyIconScale(scale) {
    const s = normalizeScale(scale);
    fabPx = Math.round(FAB_SIZE_DEFAULT * s);
    if (!rootEl) return;
    rootEl.style.setProperty("--spo-pa-fab-size", fabPx + "px");
    rootEl.style.setProperty("--spo-pa-icon-size", Math.max(20, Math.round(fabPx * 0.52)) + "px");
  }

  function ensureStyles() {
    let style = document.getElementById(STYLE_ID);
    if (!style) {
      style = document.createElement("style");
      style.id = STYLE_ID;
      (document.head || document.documentElement).appendChild(style);
    }
    style.textContent =
      "#" + ROOT_ID + "{" +
        "--spo-pa-fab-size:" + fabPx + "px;" +
        "--spo-pa-icon-size:" + Math.max(20, Math.round(fabPx * 0.52)) + "px;" +
        "--spo-pa-panel-gap:" + PANEL_GAP + "px;" +
        "position:fixed;right:" + VIEW_MARGIN + "px;bottom:" + VIEW_MARGIN + "px;left:auto;top:auto;" +
        "z-index:2147483646;font-family:'DM Sans',system-ui,'Segoe UI',sans-serif;" +
        "width:var(--spo-pa-fab-size);height:var(--spo-pa-fab-size);" +
        "box-sizing:border-box;overflow:visible !important;" +
      "}" +
      "#" + ROOT_ID + " .spo-pa-fab{" +
        "position:relative;width:var(--spo-pa-fab-size);height:var(--spo-pa-fab-size);" +
        "border-radius:50%;border:2px solid rgba(116,39,116,.65);cursor:pointer;" +
        "display:inline-flex;align-items:center;justify-content:center;padding:0;overflow:visible;" +
        "background:radial-gradient(circle at 30% 22%,#9b4d9b 0%,#742774 55%,#5a1d5a 100%);" +
        "color:#fff;box-shadow:0 12px 28px rgba(45,16,45,.35),0 0 0 1px rgba(183,148,246,.2) inset;" +
        "transition:transform .35s cubic-bezier(.33,1,.68,1),box-shadow .35s ease;" +
      "}" +
      "#" + ROOT_ID + " .spo-pa-fab::before{" +
        "content:'';position:absolute;inset:-4px;border-radius:999px;" +
        "border:2px solid rgba(183,148,246,.35);box-shadow:0 0 0 4px rgba(116,39,116,.12);" +
        "pointer-events:none;animation:spo-pa-pulse 2.8s ease-in-out infinite;" +
      "}" +
      "#" + ROOT_ID + " .spo-pa-fab:hover{" +
        "transform:translateY(-3px) scale(1.05);" +
        "box-shadow:0 18px 40px rgba(45,16,45,.45),0 0 28px rgba(116,39,116,.25);" +
      "}" +
      /* Full bubble clip — rocket launches along its nose (up-right) then returns from behind */
      "#" + ROOT_ID + " .spo-pa-fab-icon{" +
        "position:absolute;inset:0;overflow:hidden;border-radius:50%;" +
        "display:flex;align-items:center;justify-content:center;pointer-events:none;" +
      "}" +
      "#" + ROOT_ID + " .spo-pa-rocket{" +
        "width:var(--spo-pa-icon-size);height:var(--spo-pa-icon-size);display:block;" +
        "transform:translate(0,0);opacity:1;" +
      "}" +
      "#" + ROOT_ID + " .spo-pa-rocket.spo-pa-launch{" +
        "animation:spo-pa-launch .85s cubic-bezier(.45,.02,.25,1) 1;" +
      "}" +
      "@keyframes spo-pa-pulse{0%,100%{transform:scale(1);opacity:.9;}50%{transform:scale(1.05);opacity:.55;}}" +
      /* Nose points up-right (~45°); exit that way, re-enter from bottom-left */
      "@keyframes spo-pa-launch{" +
        "0%{transform:translate(0,0);opacity:1;}" +
        "40%{transform:translate(130%,-130%);opacity:0;}" +
        "41%{transform:translate(-130%,130%);opacity:0;}" +
        "100%{transform:translate(0,0);opacity:1;}" +
      "}" +
      "#" + PANEL_ID + "{" +
        "position:absolute;display:none;flex-direction:column;overflow:hidden;" +
        "width:min(92vw,1000px);min-width:min(360px,92vw);" +
        "height:70vh;min-height:280px;max-width:min(92vw,1000px);max-height:70vh;" +
        "border-radius:12px;background:#181b23;color:#e8eaef;" +
        "border:1px solid rgba(183,148,246,.22);" +
        "box-shadow:0 8px 32px rgba(0,0,0,.35);box-sizing:border-box;" +
        "font-family:'DM Sans',system-ui,'Segoe UI',sans-serif;" +
      "}" +
      "#" + ROOT_ID + ".spo-pa-open #" + PANEL_ID + "{" +
        "display:flex;animation:spo-pa-panel-fade-in " + PANEL_FADE_IN_S + "s ease-out forwards;" +
      "}" +
      "#" + ROOT_ID + ".spo-pa-closing #" + PANEL_ID + "{" +
        "display:flex;animation:spo-pa-panel-fade-out " + PANEL_FADE_OUT_S + "s ease-out forwards;" +
      "}" +
      "@keyframes spo-pa-panel-fade-in{0%{opacity:0;}100%{opacity:1;}}" +
      "@keyframes spo-pa-panel-fade-out{0%{opacity:1;}100%{opacity:0;}}" +
      "#" + ROOT_ID + ".spo-pa-panel-up #" + PANEL_ID + "{" +
        "bottom:calc(var(--spo-pa-fab-size) + var(--spo-pa-panel-gap));top:auto;" +
      "}" +
      "#" + ROOT_ID + ".spo-pa-panel-down #" + PANEL_ID + "{" +
        "top:calc(var(--spo-pa-fab-size) + var(--spo-pa-panel-gap));bottom:auto;" +
      "}" +
      "#" + ROOT_ID + ".spo-pa-anchor-left #" + PANEL_ID + "{left:0;right:auto;}" +
      "#" + ROOT_ID + ".spo-pa-anchor-right #" + PANEL_ID + "{right:0;left:auto;}" +
      "#" + PANEL_ID + " .spo-pa-header{" +
        "display:flex;flex-direction:column;align-items:stretch;gap:0;" +
        "padding:10px 14px 8px;border-bottom:1px solid rgba(255,255,255,.1);flex-shrink:0;" +
        "background:linear-gradient(135deg,#742774,#5a1d5a);" +
      "}" +
      "#" + PANEL_ID + " .spo-pa-header-left{width:100%;min-width:0;}" +
      "#" + PANEL_ID + " .spo-pa-title-row{" +
        "display:flex;align-items:center;justify-content:flex-start;gap:8px;width:100%;" +
      "}" +
      "#" + PANEL_ID + " .spo-pa-title{" +
        "flex:1;min-width:0;font-size:15px;font-weight:700;letter-spacing:.01em;line-height:1.3;" +
        "white-space:nowrap;overflow:hidden;text-overflow:ellipsis;color:#fff;" +
      "}" +
      "#" + PANEL_ID + " .spo-pa-title-name{font-weight:700;}" +
      "#" + PANEL_ID + " .spo-pa-title-sep{font-weight:400;color:rgba(255,255,255,.55);opacity:.75;}" +
      "#" + PANEL_ID + " .spo-pa-title-suffix{font-weight:400;color:rgba(255,255,255,.78);}" +
      "#" + PANEL_ID + " .spo-pa-sub{" +
        "font-size:11px;color:rgba(255,255,255,.68);margin-top:2px;line-height:1.3;" +
        "white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:100%;" +
      "}" +
      "#" + PANEL_ID + " .spo-pa-header-actions{" +
        "display:flex;align-items:center;gap:8px;flex:0 0 auto;margin-left:auto;position:relative;" +
      "}" +
      "#" + PANEL_ID + " .spo-pa-icon-btn," +
      "#" + PANEL_ID + " .spo-pa-close{" +
        "box-sizing:border-box;border:1.5px solid rgba(255,255,255,.22);" +
        "background:rgba(255,255,255,.12);color:#fff;" +
        "width:28px;height:28px;min-width:28px;border-radius:10px;cursor:pointer;" +
        "display:inline-flex;align-items:center;justify-content:center;padding:0;" +
        "flex-shrink:0;line-height:1;" +
        "transition:background .18s ease,border-color .18s ease,color .18s ease,box-shadow .18s ease;" +
      "}" +
      "#" + PANEL_ID + " .spo-pa-icon-btn:hover," +
      "#" + PANEL_ID + " .spo-pa-close:hover{" +
        "background:rgba(255,255,255,.22);color:#fff;border-color:rgba(255,255,255,.35);" +
      "}" +
      "#" + PANEL_ID + " .spo-pa-close{font-size:16px;font-weight:500;}" +
      "#" + PANEL_ID + " .spo-pa-icon-btn svg{width:17px;height:17px;display:block;opacity:.95;}" +
      /* Same Early Riser / Night Owl pill as popup + View Manager headers */
      "#" + PANEL_ID + " .spo-pa-dark-toggle{" +
        "width:52px;height:24px;border:none;border-radius:12px;cursor:pointer;" +
        "position:relative;flex-shrink:0;overflow:hidden;padding:0;" +
        "background:linear-gradient(90deg,#f5e6c8 0%,#e8d4a8 35%,#3d2a5c 65%,#1a0a2e 100%);" +
        "box-shadow:inset 0 1px 2px rgba(0,0,0,.2),0 0 0 1px rgba(255,255,255,.15);" +
      "}" +
      "#" + PANEL_ID + " .spo-pa-dark-toggle:hover{" +
        "box-shadow:inset 0 1px 2px rgba(0,0,0,.25),0 0 0 1px rgba(255,255,255,.25);" +
      "}" +
      "#" + PANEL_ID + " .spo-pa-dark-toggle::before{" +
        "content:'';position:absolute;left:5px;top:50%;transform:translateY(-50%);" +
        "width:14px;height:14px;pointer-events:none;opacity:.95;" +
        "background:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23b8860b' stroke-width='1.5' stroke-linecap='round' stroke-linejoin='round'%3E%3Ccircle cx='12' cy='10' r='4'/%3E%3Cpath d='M12 2v2M12 18v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M18 12h2M4.93 19.07l1.41-1.41M17.66 6.34l1.41-1.41'/%3E%3Cpath d='M5 14h14'/%3E%3C/svg%3E\") no-repeat center / contain;" +
      "}" +
      "#" + PANEL_ID + " .spo-pa-dark-toggle::after{" +
        "content:'';position:absolute;right:5px;top:50%;transform:translateY(-50%);" +
        "width:14px;height:14px;pointer-events:none;opacity:.9;" +
        "background:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23c9b8e8' stroke-width='1.5' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z'/%3E%3C/svg%3E\") no-repeat center / contain;" +
      "}" +
      "#" + PANEL_ID + " .spo-pa-dark-toggle-knob{" +
        "position:absolute;top:2px;left:2px;width:20px;height:20px;border-radius:50%;" +
        "background:radial-gradient(circle at 30% 30%,#fff 0%,#f0ebe0 50%,#e0d8c8 100%);" +
        "box-shadow:0 1px 3px rgba(0,0,0,.35),0 0 0 1px rgba(255,255,255,.5);" +
        "transition:transform .25s cubic-bezier(.4,0,.2,1);pointer-events:none;" +
        "display:flex;align-items:center;justify-content:center;" +
      "}" +
      "#" + PANEL_ID + " .spo-pa-dark-toggle-knob::before{" +
        "content:'';width:11px;height:11px;" +
        "background:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23c4952a' stroke-width='1.8' stroke-linecap='round' stroke-linejoin='round'%3E%3Ccircle cx='12' cy='9' r='3'/%3E%3Cpath d='M12 1v1.5M12 19.5V21M4.22 4.22l1.06 1.06M18.72 18.72l1.06 1.06M1 12h1.5M19.5 12H21M4.22 19.78l1.06-1.06M18.72 5.28l1.06-1.06'/%3E%3Cpath d='M4 14h16'/%3E%3C/svg%3E\") no-repeat center / contain;" +
      "}" +
      "#" + PANEL_ID + " .spo-pa-dark-toggle.is-night .spo-pa-dark-toggle-knob{" +
        "transform:translateX(28px);" +
        "background:radial-gradient(circle at 30% 30%,#c9b8e8 0%,#5c4d7a 50%,#2d2345 100%);" +
        "box-shadow:0 1px 3px rgba(0,0,0,.5),0 0 0 1px rgba(255,255,255,.15);" +
      "}" +
      "#" + PANEL_ID + " .spo-pa-dark-toggle.is-night .spo-pa-dark-toggle-knob::before{" +
        "width:10px;height:10px;" +
        "background:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23e8e0f0' stroke-width='1.5' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z'/%3E%3C/svg%3E\") no-repeat center / contain;" +
      "}" +
      "#" + PANEL_ID + " .spo-pa-mode-menu{" +
        "display:none;position:absolute;top:calc(100% + 6px);right:0;z-index:5;" +
        "min-width:200px;padding:6px;border-radius:9px;background:#181b23;color:#e8eaef;" +
        "border:1px solid rgba(183,148,246,.28);box-shadow:0 10px 28px rgba(0,0,0,.4);" +
        "font-family:'DM Sans',system-ui,'Segoe UI',sans-serif;" +
      "}" +
      "#" + PANEL_ID + " .spo-pa-mode-menu.is-open{display:block;}" +
      "#" + PANEL_ID + " .spo-pa-mode-menu-label{" +
        "font-size:10px;font-weight:700;letter-spacing:.04em;text-transform:uppercase;" +
        "color:rgba(232,234,239,.55);padding:6px 8px 4px;" +
      "}" +
      "#" + PANEL_ID + " .spo-pa-mode-item{" +
        "display:block;width:100%;border:none;background:transparent;color:inherit;" +
        "text-align:left;padding:8px;border-radius:7px;cursor:pointer;font-size:12px;" +
      "}" +
      "#" + PANEL_ID + " .spo-pa-mode-item:hover{background:rgba(183,148,246,.12);}" +
      "#" + PANEL_ID + " .spo-pa-mode-item.is-active{background:rgba(183,148,246,.2);}" +
      "#" + PANEL_ID + " .spo-pa-mode-item small{display:block;opacity:.65;margin-top:2px;font-size:11px;}" +
      "#" + PANEL_ID + " .spo-pa-body{flex:1;min-height:0;display:flex;flex-direction:column;}" +
      "#" + FRAME_ID + "{flex:1;width:100%;height:100%;border:0;background:#181b23;min-height:0;}" +
      "@media (max-width:720px){" +
        "#" + PANEL_ID + "{" +
          "width:min(96vw,1000px);min-width:0;height:min(78vh,1000px);max-height:78vh;" +
        "}" +
      "}";
  }

  function rocketSvg() {
    return (
      '<svg class="spo-pa-rocket" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" ' +
      'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
      '<path d="M4.5 16.5c-1.5 1.26-2 5-2 5s3.74-.5 5-2c.71-.84.7-2.13-.09-2.91a2.18 2.18 0 0 0-2.91-.09z"/>' +
      '<path d="m12 15-3-3a22 22 0 0 1 2-3.95A12.88 12.88 0 0 1 22 2c0 2.72-.78 7.5-6 11a22.35 22.35 0 0 1-4 2z"/>' +
      '<path d="M9 12H4s.55-3.03 2-4c1.62-1.08 5 0 5 0"/>' +
      '<path d="M12 15v5s3.03-.55 4-2c1.08-1.62 0-5 0-5"/>' +
      "</svg>"
    );
  }

  function buildLauncherMarkup() {
    return (
      '<button type="button" class="spo-pa-fab" title="SPO Dev Toolkit — Power Apps" aria-label="Open Power Apps toolkit popout" aria-expanded="false">' +
      '<span class="spo-pa-fab-icon">' +
      rocketSvg() +
      "</span></button>" +
      '<div id="' +
      PANEL_ID +
      '" role="dialog" aria-label="Power Apps toolkit">' +
      '<div class="spo-pa-header"><div class="spo-pa-header-left">' +
      '<div class="spo-pa-title-row">' +
      '<div class="spo-pa-title" id="spo-pa-title">' +
      '<span class="spo-pa-title-name">Power Apps</span>' +
      '<span class="spo-pa-title-sep" aria-hidden="true"> | </span>' +
      '<span class="spo-pa-title-suffix">Tools</span></div>' +
      '<div class="spo-pa-header-actions">' +
      '<button type="button" class="spo-pa-icon-btn" id="spo-pa-display-mode" title="Display mode" aria-label="Display mode" aria-haspopup="menu" aria-expanded="false">' +
      GEAR_SVG +
      "</button>" +
      '<div class="spo-pa-mode-menu" id="spo-pa-mode-menu" role="menu">' +
      '<div class="spo-pa-mode-menu-label">Display Mode</div>' +
      '<button type="button" class="spo-pa-mode-item" data-mode="default" role="menuitem">' +
      "Default<small>All features visible</small></button>" +
      '<button type="button" class="spo-pa-mode-item" data-mode="simple" role="menuitem">' +
      "Simple<small>Essential features only</small></button>" +
      "</div>" +
      '<button type="button" class="spo-pa-dark-toggle" id="spo-pa-dark-toggle" title="Early Riser" aria-label="Theme: Early Riser. Click to switch to Night Owl.">' +
      '<span class="spo-pa-dark-toggle-knob" aria-hidden="true"></span></button>' +
      '<button type="button" class="spo-pa-close" title="Close" aria-label="Close">×</button>' +
      "</div></div>" +
      '<div class="spo-pa-sub" id="spo-pa-sub">Power Apps / Dynamics</div>' +
      "</div></div>" +
      '<div class="spo-pa-body">' +
      '<iframe id="' +
      FRAME_ID +
      '" title="Power Apps toolkit" src="about:blank"></iframe>' +
      "</div></div>"
    );
  }

  function resolveIsDark(storage) {
    if (typeof storage?.darkMode === "boolean") return storage.darkMode;
    if (storage?.[THEME_KEY] === "dark" || storage?.[THEME_KEY] === "light") {
      return storage[THEME_KEY] === "dark";
    }
    return false;
  }

  function updatePanelPlacement() {
    if (!rootEl) return;
    const rect = rootEl.getBoundingClientRect();
    const vw = window.innerWidth;
    const vh = window.innerHeight;
    const cx = rect.left + rect.width / 2;
    const cy = rect.top + rect.height / 2;
    const openDown = cy < vh / 2;
    const anchorLeft = cx < vw / 2;
    rootEl.classList.remove(
      "spo-pa-panel-up",
      "spo-pa-panel-down",
      "spo-pa-anchor-left",
      "spo-pa-anchor-right"
    );
    rootEl.classList.add(openDown ? "spo-pa-panel-down" : "spo-pa-panel-up");
    rootEl.classList.add(anchorLeft ? "spo-pa-anchor-left" : "spo-pa-anchor-right");
  }

  function clampPosition(left, top) {
    const maxL = Math.max(VIEW_MARGIN, window.innerWidth - fabPx - VIEW_MARGIN);
    const maxT = Math.max(VIEW_MARGIN, window.innerHeight - fabPx - VIEW_MARGIN);
    return {
      left: Math.min(maxL, Math.max(VIEW_MARGIN, left)),
      top: Math.min(maxT, Math.max(VIEW_MARGIN, top)),
    };
  }

  function applyPosition(storedPos) {
    if (!rootEl) return;
    if (
      storedPos &&
      typeof storedPos.left === "number" &&
      typeof storedPos.top === "number" &&
      !Number.isNaN(storedPos.left) &&
      !Number.isNaN(storedPos.top)
    ) {
      const c = clampPosition(storedPos.left, storedPos.top);
      rootEl.style.left = c.left + "px";
      rootEl.style.top = c.top + "px";
      rootEl.style.right = "auto";
      rootEl.style.bottom = "auto";
    } else {
      rootEl.style.left = "auto";
      rootEl.style.top = "auto";
      rootEl.style.right = VIEW_MARGIN + "px";
      rootEl.style.bottom = VIEW_MARGIN + "px";
    }
    updatePanelPlacement();
  }

  function persistPosition(left, top) {
    const c = clampPosition(left, top);
    try {
      chrome.storage.local.set({ [POSITION_KEY]: { left: c.left, top: c.top } });
    } catch (_) {}
  }

  function loadSharedLauncherPrefs() {
    chrome.storage.local.get([POSITION_KEY, ICON_SCALE_KEY, LEGACY_POSITION_KEY], (r) => {
      applyIconScale(r[ICON_SCALE_KEY]);
      ensureStyles();
      let pos = r[POSITION_KEY];
      if (!pos && r[LEGACY_POSITION_KEY]) {
        const legacy = r[LEGACY_POSITION_KEY];
        if (typeof legacy.left === "number" && typeof legacy.top === "number") {
          pos = { left: legacy.left, top: legacy.top };
        } else if (typeof legacy.right === "number" && typeof legacy.bottom === "number") {
          pos = {
            left: Math.max(VIEW_MARGIN, window.innerWidth - fabPx - legacy.right),
            top: Math.max(VIEW_MARGIN, window.innerHeight - fabPx - legacy.bottom),
          };
        }
        if (pos) chrome.storage.local.set({ [POSITION_KEY]: pos });
      }
      applyPosition(pos);
    });
  }

  function playRocketLaunch() {
    if (!rootEl) return;
    const rocket = rootEl.querySelector(".spo-pa-rocket");
    if (!rocket) return;
    rocket.classList.remove("spo-pa-launch");
    // Restart CSS animation if already mid-flight
    void rocket.offsetWidth;
    rocket.classList.add("spo-pa-launch");
  }

  function escapeHtml(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function applyHeaderContext(ctx) {
    if (!rootEl) return;
    const titleEl = rootEl.querySelector("#spo-pa-title");
    const subEl = rootEl.querySelector("#spo-pa-sub");
    const fab = rootEl.querySelector(".spo-pa-fab");
    const appName = String((ctx && ctx.displayName) || "").trim() || "Power Apps";
    const orgLabel = String((ctx && ctx.orgLabel) || "").trim();
    let clientUrl = String((ctx && ctx.clientUrl) || "").trim();
    let host = "";
    try {
      if (clientUrl) host = new URL(clientUrl).hostname;
    } catch (_) {}
    if (!clientUrl) {
      try {
        clientUrl = window.location.origin;
        host = window.location.hostname;
      } catch (_) {}
    }

    if (titleEl) {
      titleEl.innerHTML =
        '<span class="spo-pa-title-name">' +
        escapeHtml(appName) +
        "</span>" +
        '<span class="spo-pa-title-sep" aria-hidden="true"> | </span>' +
        '<span class="spo-pa-title-suffix">Tools</span>';
    }
    if (subEl) {
      if (orgLabel && clientUrl) {
        subEl.textContent = orgLabel + " - " + clientUrl;
      } else if (clientUrl) {
        subEl.textContent = (host ? host + " - " : "") + clientUrl;
      } else {
        subEl.textContent = "Power Apps / Dynamics";
      }
      if (clientUrl) subEl.title = clientUrl;
    }
    if (fab) {
      fab.title = appName + " — Tools";
      fab.setAttribute("aria-label", "Open toolkit for " + appName);
    }
  }

  function refreshHeaderContext() {
    chrome.runtime.sendMessage({ type: "SPOToolkitGetPowerAppsHeaderContext" }, (res) => {
      if (chrome.runtime.lastError || !res || !res.ok) {
        applyHeaderContext({
          displayName: "",
          clientUrl: window.location.origin,
          orgLabel: "",
        });
        return;
      }
      applyHeaderContext(res.context || {});
    });
  }

  function ensureSidebarFrame() {
    const frame = document.getElementById(FRAME_ID);
    if (!frame) return;
    const want = chrome.runtime.getURL("levelup/sidebar.html") + "?host=popout";
    if (!frame.src || frame.src === "about:blank" || frame.src.indexOf("host=popout") === -1) {
      frame.src = want;
    }
  }

  function setOpen(open) {
    if (!rootEl) return;
    const wantOpen = !!open;
    const btn = rootEl.querySelector(".spo-pa-fab");
    if (wantOpen) {
      if (closeAnimTimer) {
        clearTimeout(closeAnimTimer);
        closeAnimTimer = null;
      }
      expanded = true;
      rootEl.classList.remove("spo-pa-closing");
      rootEl.classList.add("spo-pa-open");
      if (btn) btn.setAttribute("aria-expanded", "true");
      updatePanelPlacement();
      refreshHeaderContext();
      ensureSidebarFrame();
      return;
    }
    if (!expanded && !rootEl.classList.contains("spo-pa-open")) return;
    expanded = false;
    if (btn) btn.setAttribute("aria-expanded", "false");
    rootEl.classList.remove("spo-pa-open");
    rootEl.classList.add("spo-pa-closing");
    if (closeAnimTimer) clearTimeout(closeAnimTimer);
    closeAnimTimer = setTimeout(() => {
      if (!rootEl) return;
      rootEl.classList.remove("spo-pa-closing");
      closeAnimTimer = null;
    }, PANEL_FADE_OUT_MS);
  }

  function bindDrag(fab) {
    let dragging = false;
    let startX = 0;
    let startY = 0;
    let startLeft = 0;
    let startTop = 0;
    let moved = false;

    fab.addEventListener("pointerdown", (e) => {
      if (e.button != null && e.button !== 0) return;
      dragging = true;
      moved = false;
      startX = e.clientX;
      startY = e.clientY;
      const rect = rootEl.getBoundingClientRect();
      startLeft = rect.left;
      startTop = rect.top;
      rootEl.style.left = startLeft + "px";
      rootEl.style.top = startTop + "px";
      rootEl.style.right = "auto";
      rootEl.style.bottom = "auto";
      fab.setPointerCapture(e.pointerId);
    });

    fab.addEventListener("pointermove", (e) => {
      if (!dragging) return;
      const dx = e.clientX - startX;
      const dy = e.clientY - startY;
      if (Math.abs(dx) + Math.abs(dy) > 4) moved = true;
      const c = clampPosition(startLeft + dx, startTop + dy);
      rootEl.style.left = c.left + "px";
      rootEl.style.top = c.top + "px";
      if (expanded) updatePanelPlacement();
    });

    fab.addEventListener("pointerup", (e) => {
      if (!dragging) return;
      dragging = false;
      try {
        fab.releasePointerCapture(e.pointerId);
      } catch (_) {}
      if (moved) {
        const left = parseFloat(rootEl.style.left) || VIEW_MARGIN;
        const top = parseFloat(rootEl.style.top) || VIEW_MARGIN;
        persistPosition(left, top);
        updatePanelPlacement();
      } else {
        playRocketLaunch();
        setOpen(!expanded);
      }
    });
  }

  function onViewportChange() {
    if (resizeTimer) clearTimeout(resizeTimer);
    resizeTimer = setTimeout(() => {
      chrome.storage.local.get(POSITION_KEY, (r) => {
        applyPosition(r[POSITION_KEY]);
      });
    }, 120);
  }

  function mount() {
    if (document.getElementById(ROOT_ID)) {
      rootEl = document.getElementById(ROOT_ID);
      return;
    }
    ensureStyles();
    rootEl = document.createElement("div");
    rootEl.id = ROOT_ID;
    rootEl.classList.add("spo-pa-panel-up", "spo-pa-anchor-right");
    rootEl.innerHTML = buildLauncherMarkup();

    document.documentElement.appendChild(rootEl);
    const fab = rootEl.querySelector(".spo-pa-fab");
    const closeBtn = rootEl.querySelector(".spo-pa-close");
    const rocket = rootEl.querySelector(".spo-pa-rocket");
    bindDrag(fab);
    if (rocket) {
      rocket.addEventListener("animationend", (e) => {
        if (e.animationName === "spo-pa-launch") {
          rocket.classList.remove("spo-pa-launch");
        }
      });
    }
    closeBtn.addEventListener("click", () => setOpen(false));
    bindHeaderChrome();
    refreshHeaderContext();
    loadSharedLauncherPrefs();
    window.addEventListener("resize", onViewportChange);
  }

  function applyDarkToggleUi(isDark) {
    if (!rootEl) return;
    const btn = rootEl.querySelector("#spo-pa-dark-toggle");
    if (!btn) return;
    btn.classList.toggle("is-night", !!isDark);
    if (isDark) {
      btn.title = "Night Owl";
      btn.setAttribute(
        "aria-label",
        "Theme: Night Owl. Click to switch to Early Riser."
      );
    } else {
      btn.title = "Early Riser";
      btn.setAttribute(
        "aria-label",
        "Theme: Early Riser. Click to switch to Night Owl."
      );
    }
  }

  function applyDisplayModeMenuUi(mode) {
    if (!rootEl) return;
    rootEl.querySelectorAll(".spo-pa-mode-item").forEach((item) => {
      item.classList.toggle("is-active", item.getAttribute("data-mode") === mode);
    });
  }

  function setModeMenuOpen(open) {
    if (!rootEl) return;
    const menu = rootEl.querySelector("#spo-pa-mode-menu");
    const btn = rootEl.querySelector("#spo-pa-display-mode");
    if (menu) menu.classList.toggle("is-open", !!open);
    if (btn) btn.setAttribute("aria-expanded", open ? "true" : "false");
  }

  function bindHeaderChrome() {
    if (!rootEl) return;
    const darkBtn = rootEl.querySelector("#spo-pa-dark-toggle");
    const modeBtn = rootEl.querySelector("#spo-pa-display-mode");
    const modeMenu = rootEl.querySelector("#spo-pa-mode-menu");

    chrome.storage.local.get(["darkMode", THEME_KEY, CONFIG_KEY], (r) => {
      applyDarkToggleUi(resolveIsDark(r));
      const mode = r[CONFIG_KEY]?.displayMode === "simple" ? "simple" : "default";
      applyDisplayModeMenuUi(mode);
    });

    if (darkBtn) {
      darkBtn.addEventListener("click", () => {
        chrome.storage.local.get(["darkMode", THEME_KEY], (r) => {
          const next = !resolveIsDark(r);
          chrome.storage.local.set({
            darkMode: next,
            [THEME_KEY]: next ? "dark" : "light",
          });
          applyDarkToggleUi(next);
        });
      });
    }

    if (modeBtn && modeMenu) {
      modeBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        setModeMenuOpen(!modeMenu.classList.contains("is-open"));
      });
      modeMenu.addEventListener("click", (e) => {
        const item = e.target.closest(".spo-pa-mode-item");
        if (!item) return;
        const mode = item.getAttribute("data-mode") === "simple" ? "simple" : "default";
        chrome.storage.local.get(CONFIG_KEY, (r) => {
          const current = r[CONFIG_KEY] && typeof r[CONFIG_KEY] === "object" ? r[CONFIG_KEY] : {};
          const next = Object.assign({}, current, DISPLAY_MODE_PRESETS[mode], {
            displayMode: mode,
          });
          chrome.storage.local.set({ [CONFIG_KEY]: next });
          applyDisplayModeMenuUi(mode);
          setModeMenuOpen(false);
        });
      });
      document.addEventListener("click", (e) => {
        if (!rootEl) return;
        if (modeMenu.contains(e.target) || modeBtn.contains(e.target)) return;
        setModeMenuOpen(false);
      });
    }

    if (chrome.storage?.onChanged) {
      chrome.storage.onChanged.addListener((changes, area) => {
        if (area !== "local") return;
        if (changes.darkMode || changes[THEME_KEY]) {
          const isDark =
            typeof changes.darkMode?.newValue === "boolean"
              ? changes.darkMode.newValue
              : changes[THEME_KEY]?.newValue === "dark";
          if (typeof changes.darkMode?.newValue === "boolean" || changes[THEME_KEY]) {
            applyDarkToggleUi(!!isDark);
          }
        }
        if (changes[CONFIG_KEY]?.newValue?.displayMode) {
          applyDisplayModeMenuUi(changes[CONFIG_KEY].newValue.displayMode);
        }
      });
    }
  }

  function openPopout() {
    if (!isDynamicsRuntime()) {
      return { ok: false, error: "Not on a model-driven app runtime" };
    }
    mount();
    setOpen(true);
    return { ok: true };
  }

  function closePopout() {
    setOpen(false);
    return { ok: true };
  }

  chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
    if (!message || !message.type) return;
    if (message.type === "SPOToolkitOpenPowerAppsPopout") {
      sendResponse(openPopout());
      return false;
    }
    if (message.type === "SPOToolkitClosePowerAppsPopout") {
      sendResponse(closePopout());
      return false;
    }
    if (message.type === "SPOToolkitPowerAppsPopoutPing") {
      sendResponse({ ok: true, runtime: isDynamicsRuntime(), open: expanded });
      return false;
    }
  });

  chrome.storage.onChanged.addListener((changes, area) => {
    if (area !== "local" || !rootEl) return;
    if (changes[POSITION_KEY]) {
      applyPosition(changes[POSITION_KEY].newValue);
    }
    if (changes[ICON_SCALE_KEY]) {
      applyIconScale(changes[ICON_SCALE_KEY].newValue);
      ensureStyles();
      chrome.storage.local.get(POSITION_KEY, (r) => applyPosition(r[POSITION_KEY]));
    }
  });

  function maybeMountFab() {
    if (!isDynamicsRuntime()) return;
    mount();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => setTimeout(maybeMountFab, 800));
  } else {
    setTimeout(maybeMountFab, 800);
  }
  setInterval(() => {
    if (isDynamicsRuntime() && !document.getElementById(ROOT_ID)) maybeMountFab();
  }, 3000);
})();
