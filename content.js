// Runs on SharePoint pages. Injects scripts into page context so fetch uses session.

const PROGRESS_BOX_ID = "sp-csv-export-progress";
const SCRIPT_TIMEOUT = 15000;
/** JSON script node read by getViewsData.js (page context). */
const SP_VIEWS_PARAMS_SCRIPT_ID = "sp-views-params";

/** Injects/replaces #sp-views-params for View Manager REST (listId, optional viewId, webAbsoluteUrl). */
function attachSpViewsParamsScript(message) {
  const payload = {
    viewId: message.viewId || null,
    listId: message.listId || null,
    webAbsoluteUrl: message.webAbsoluteUrl || null,
  };
  let el = document.getElementById(SP_VIEWS_PARAMS_SCRIPT_ID);
  if (el) el.remove();
  el = document.createElement("script");
  el.id = SP_VIEWS_PARAMS_SCRIPT_ID;
  el.type = "application/json";
  el.textContent = JSON.stringify(payload);
  (document.head || document.documentElement).appendChild(el);
}

function showProgress(message) {
  let el = document.getElementById(PROGRESS_BOX_ID);
  if (!el) {
    el = document.createElement("div");
    el.id = PROGRESS_BOX_ID;
    el.style.cssText =
      "position:fixed;bottom:20px;right:20px;z-index:2147483647;font-family:'Segoe UI',sans-serif;min-width:280px;max-width:90vw;";
    const wrap = document.createElement("div");
    wrap.className = "sp-csv-progress-wrap";
    wrap.innerHTML =
      '<div class="sp-csv-progress-rotate"></div>' +
      '<div class="sp-csv-progress-mask"></div>' +
      '<div class="sp-csv-progress-content">' +
      '<div class="sp-csv-progress-spinner"></div>' +
      '<span class="sp-csv-msg">' + (message || "Starting…") + "</span>" +
      "</div>";
    el.appendChild(wrap);
    const style = document.createElement("style");
    style.textContent =
      ".sp-csv-progress-wrap{position:relative;border-radius:11px;overflow:hidden;}" +
      ".sp-csv-progress-rotate{position:absolute;left:-50%;top:-50%;width:200%;height:200%;background:conic-gradient(transparent,#37ae1c 12%,#3D4FD6 28%,transparent 42%);animation:spcsvrotate 4s linear infinite;z-index:0;}" +
      ".sp-csv-progress-wrap.dark-mode .sp-csv-progress-rotate{background:conic-gradient(transparent,#86efac 12%,#7B89F5 28%,transparent 42%);}" +
      ".sp-csv-progress-mask{position:absolute;left:6px;top:6px;width:calc(100% - 12px);height:calc(100% - 12px);background:#323130;border-radius:5px;z-index:1;}" +
      ".sp-csv-progress-content{position:relative;z-index:2;display:flex;align-items:center;gap:12px;padding:14px 20px;color:#fff;font-size:14px;}" +
      ".sp-csv-progress-spinner{width:20px;height:20px;border:2px solid rgba(255,255,255,0.3);border-top-color:#fff;border-radius:50%;animation:spcsvspin 0.8s linear infinite;flex-shrink:0;}" +
      "@keyframes spcsvrotate{to{transform:rotate(1turn);}}" +
      "@keyframes spcsvspin{to{transform:rotate(360deg);}}";
    document.head.appendChild(style);
    document.body.appendChild(el);
    chrome.storage.local.get("darkMode", (r) => {
      if (r.darkMode && wrap) wrap.classList.add("dark-mode");
    });
  }
  const msgEl = el.querySelector(".sp-csv-msg");
  if (msgEl) msgEl.textContent = message || "…";
}

function finishProgress(success, message, stopReason) {
  const el = document.getElementById(PROGRESS_BOX_ID);
  if (!el) return;
  const rotate = el.querySelector(".sp-csv-progress-rotate");
  const mask = el.querySelector(".sp-csv-progress-mask");
  const msgEl = el.querySelector(".sp-csv-msg");
  const spinner = el.querySelector(".sp-csv-progress-spinner");
  if (rotate) {
    rotate.style.animation = "none";
    rotate.style.background = success ? "#107c10" : "#a4262c";
  }
  if (mask) mask.style.background = success ? "#107c10" : "#a4262c";
  if (spinner) spinner.style.animation = "none";
  if (msgEl) msgEl.textContent = message || (success ? "Done!" : "Error.");
  const hasEarlyStop = stopReason && /no View ID|returned 0 rows/.test(stopReason);
  const displayMs = (success && stopReason) ? 15000 : (success ? 3500 : 10000);
  if (hasEarlyStop) alert("SharePoint CSV Export: " + stopReason);
  setTimeout(() => {
    const e = document.getElementById(PROGRESS_BOX_ID);
    if (e && e.parentNode) e.parentNode.removeChild(e);
  }, displayMs);
}

const FACT_TOAST_ID = "sp-fact-toast";
const FACT_TOAST_STYLE_ID = "sp-fact-toast-styles";
const FACT_TOAST_DURATION_MS = 45000;
let factToastTimeoutId = null;

const FACT_TOAST_CSS =
  ".sp-fact-wrap{position:relative;background:#323130;color:#fff;border-radius:11px;padding:14px 36px 14px 16px;box-shadow:0 4px 24px rgba(0,0,0,.5);font-size:13px;line-height:1.5;}" +
  ".sp-fact-text{margin:0 0 8px;}" +
  ".sp-fact-source{display:inline-block;padding:6px 12px;background:rgba(255,255,255,.2);color:#fff;font-size:12px;font-weight:600;text-decoration:none;border-radius:6px;}" +
  ".sp-fact-source:hover{background:rgba(255,255,255,.3);}" +
  ".sp-fact-close{position:absolute;top:8px;right:8px;width:24px;height:24px;border:none;background:rgba(255,255,255,.15);color:#fff;border-radius:4px;font-size:18px;line-height:1;cursor:pointer;display:flex;align-items:center;justify-content:center;padding:0;}" +
  ".sp-fact-close:hover{background:rgba(255,255,255,.25);}";

function showFactToast(factText, sourceUrl) {
  if (factToastTimeoutId) {
    clearTimeout(factToastTimeoutId);
    factToastTimeoutId = null;
  }
  let el = document.getElementById(FACT_TOAST_ID);
  if (el) el.remove();
  if (!document.getElementById(FACT_TOAST_STYLE_ID)) {
    const style = document.createElement("style");
    style.id = FACT_TOAST_STYLE_ID;
    style.textContent = FACT_TOAST_CSS;
    document.head.appendChild(style);
  }
  el = document.createElement("div");
  el.id = FACT_TOAST_ID;
  el.style.cssText = "position:fixed;bottom:20px;right:20px;z-index:2147483647;font-family:'Segoe UI',sans-serif;max-width:320px;min-width:200px;";
  el.innerHTML =
    '<div class="sp-fact-wrap">' +
    '<button type="button" class="sp-fact-close" aria-label="Close">×</button>' +
    '<p class="sp-fact-text"></p>' +
    (sourceUrl ? '<a class="sp-fact-source" href="#" target="_blank" rel="noopener">Learn More</a>' : '') +
    '</div>';
  el.querySelector(".sp-fact-text").textContent = factText || "";
  if (sourceUrl) {
    const link = el.querySelector(".sp-fact-source");
    link.href = sourceUrl;
    link.addEventListener("click", (e) => { e.preventDefault(); window.open(sourceUrl, "_blank", "noopener"); });
  }
  document.body.appendChild(el);
  const close = () => {
    factToastTimeoutId = null;
    const node = document.getElementById(FACT_TOAST_ID);
    if (node && node.parentNode) node.parentNode.removeChild(node);
  };
  el.querySelector(".sp-fact-close").addEventListener("click", close);
  factToastTimeoutId = setTimeout(close, FACT_TOAST_DURATION_MS);
}

// --- Lists/Libraries launcher (floating button, bottom-right) ---
const LISTS_LAUNCHER_ID = "sp-toolkit-lists-launcher";
const LISTS_LAUNCHER_PANEL_ID = "sp-toolkit-lists-panel";
const LISTS_LAUNCHER_STYLE_ID = "sp-toolkit-lists-launcher-style";

const LISTS_LAUNCHER_CSS =
  "#" + LISTS_LAUNCHER_ID + "{--sp-toolkit-launcher-panel-gap:12px;--sp-toolkit-launcher-hover-dur:.52s;--sp-toolkit-launcher-hover-ease:cubic-bezier(0.33,1,0.68,1);position:fixed;right:14px;bottom:14px;left:auto;top:auto;z-index:2147483646;font-family:'DM Sans',system-ui,'Segoe UI',sans-serif;transition:opacity .15s ease-out;display:block;width:var(--sp-toolkit-launcher-size,58px);flex-shrink:0;box-sizing:border-box;touch-action:manipulation;}" +
  "#" + LISTS_LAUNCHER_ID + ".sp-toolkit-launcher-dragging .sp-toolkit-lists-btn{cursor:grabbing;touch-action:none;user-select:none;-webkit-user-select:none;transition:none!important;transform:none!important;box-shadow:0 6px 14px rgba(0,0,0,.25)!important;}" +
  "#" + LISTS_LAUNCHER_ID + ".sp-toolkit-launcher-dragging .sp-toolkit-lists-btn::before{animation:none!important;opacity:.35!important;box-shadow:none!important;}" +
  "#" + LISTS_LAUNCHER_ID + ".sp-toolkit-launcher-wiggle .sp-toolkit-lists-btn{animation:sp-toolkit-launcher-wiggle .12s ease-in-out infinite;}" +
  "#" + LISTS_LAUNCHER_ID + ".sp-toolkit-launcher-wiggle .sp-toolkit-lists-btn:hover{transform:none;}" +
  "@keyframes sp-toolkit-launcher-wiggle{0%,100%{transform:rotate(-2.2deg);}50%{transform:rotate(2.2deg);}}" +
  "#" + LISTS_LAUNCHER_ID + ".sp-toolkit-launcher-rolled-up{opacity:0;pointer-events:none;}" +
  "#" + LISTS_LAUNCHER_ID + " .sp-toolkit-lists-btn{position:relative;width:var(--sp-toolkit-launcher-size,58px)!important;height:var(--sp-toolkit-launcher-size,58px)!important;min-width:var(--sp-toolkit-launcher-size,58px)!important;min-height:var(--sp-toolkit-launcher-size,58px)!important;max-width:var(--sp-toolkit-launcher-size,58px)!important;max-height:var(--sp-toolkit-launcher-size,58px)!important;flex:0 0 var(--sp-toolkit-launcher-size,58px)!important;border-radius:50%;border:2px solid rgba(61,79,214,.58);cursor:pointer;display:inline-flex!important;align-items:center;justify-content:center;padding:0!important;box-sizing:border-box!important;overflow:visible;" +
  "background:radial-gradient(circle at 30% 25%,#ffffff 0%,#eef3ff 62%,#dde7ff 100%);color:#1a1d24;box-shadow:0 12px 28px rgba(22,34,66,.28),0 0 0 1px rgba(61,79,214,.15) inset;transition:transform var(--sp-toolkit-launcher-hover-dur,.52s) var(--sp-toolkit-launcher-hover-ease,cubic-bezier(0.33,1,0.68,1)),box-shadow var(--sp-toolkit-launcher-hover-dur,.52s) var(--sp-toolkit-launcher-hover-ease,cubic-bezier(0.33,1,0.68,1)),border-color var(--sp-toolkit-launcher-hover-dur,.52s) var(--sp-toolkit-launcher-hover-ease,cubic-bezier(0.33,1,0.68,1)),filter .58s var(--sp-toolkit-launcher-hover-ease,cubic-bezier(0.33,1,0.68,1));}" +
  "#" + LISTS_LAUNCHER_ID + " .sp-toolkit-lists-btn::before{content:'';position:absolute;inset:-6px;border-radius:999px;border:2px solid rgba(61,79,214,.28);box-shadow:0 0 0 6px rgba(61,79,214,.12);opacity:.95;pointer-events:none;animation:sp-toolkit-launcher-pulse 2.8s ease-in-out infinite;transition:opacity var(--sp-toolkit-launcher-hover-dur,.52s) var(--sp-toolkit-launcher-hover-ease,cubic-bezier(0.33,1,0.68,1)),transform var(--sp-toolkit-launcher-hover-dur,.52s) var(--sp-toolkit-launcher-hover-ease,cubic-bezier(0.33,1,0.68,1)),border-color var(--sp-toolkit-launcher-hover-dur,.52s) var(--sp-toolkit-launcher-hover-ease,cubic-bezier(0.33,1,0.68,1)),box-shadow var(--sp-toolkit-launcher-hover-dur,.52s) var(--sp-toolkit-launcher-hover-ease,cubic-bezier(0.33,1,0.68,1));}" +
  "#" + LISTS_LAUNCHER_ID + " .sp-toolkit-lists-btn:hover{transform:translateY(-4px) scale(1.085);border-color:rgba(61,79,214,.92);box-shadow:0 22px 44px rgba(22,34,66,.38),0 0 0 1px rgba(61,79,214,.28) inset,0 0 28px rgba(61,79,214,.18);filter:saturate(1.12) brightness(1.03);will-change:transform;}" +
  "#" + LISTS_LAUNCHER_ID + " .sp-toolkit-lists-btn:hover::before{opacity:.78;transform:scale(1.06);animation-play-state:paused;}" +
  "#" + LISTS_LAUNCHER_ID + ".sp-toolkit-launcher-dark .sp-toolkit-lists-btn{background:radial-gradient(circle at 30% 22%,#3f4656 0%,#273142 62%,#1f2633 100%);border:2px solid rgba(55,174,28,.58);box-shadow:0 14px 30px rgba(0,0,0,.5),0 0 0 1px rgba(55,174,28,.16) inset;}" +
  "#" + LISTS_LAUNCHER_ID + ".sp-toolkit-launcher-dark .sp-toolkit-lists-btn::before{border-color:rgba(134,239,172,.34);box-shadow:0 0 0 6px rgba(55,174,28,.14);}" +
  "#" + LISTS_LAUNCHER_ID + ".sp-toolkit-launcher-dark .sp-toolkit-lists-btn:hover{border-color:rgba(134,239,172,.82);box-shadow:0 22px 48px rgba(0,0,0,.58),0 0 0 1px rgba(134,239,172,.24) inset,0 0 32px rgba(55,174,28,.2);filter:saturate(1.08) brightness(1.04);}" +
  "#" + LISTS_LAUNCHER_ID + ".sp-toolkit-launcher-dark .sp-toolkit-lists-btn .sp-toolkit-launcher-compass{filter:brightness(0) invert(1);}" +
  "#" + LISTS_LAUNCHER_ID + " .sp-toolkit-lists-btn .sp-toolkit-launcher-compass,#" + LISTS_LAUNCHER_ID + " .sp-toolkit-lists-btn img.sp-toolkit-launcher-compass{width:var(--sp-toolkit-launcher-icon-size,30px)!important;height:var(--sp-toolkit-launcher-icon-size,30px)!important;min-width:var(--sp-toolkit-launcher-icon-size,30px)!important;min-height:var(--sp-toolkit-launcher-icon-size,30px)!important;max-width:var(--sp-toolkit-launcher-icon-size,30px)!important;max-height:var(--sp-toolkit-launcher-icon-size,30px)!important;display:block;flex-shrink:0;object-fit:contain;transition:transform var(--sp-toolkit-launcher-hover-dur,.52s) var(--sp-toolkit-launcher-hover-ease,cubic-bezier(0.33,1,0.68,1));-webkit-user-drag:none;user-drag:none;}" +
  "@keyframes sp-toolkit-launcher-pulse{0%,100%{transform:scale(1);opacity:.92;}50%{transform:scale(1.045);opacity:.55;}}" +
  "@keyframes sp-toolkit-launcher-tape-cw{to{transform:rotate(360deg);}}" +
  "@keyframes sp-toolkit-launcher-tape-ccw{to{transform:rotate(-360deg);}}" +
  "#" + LISTS_LAUNCHER_ID + " .sp-toolkit-launcher-compass.sp-toolkit-launcher-compass-tape-cw{animation:sp-toolkit-launcher-tape-cw .58s cubic-bezier(.45,.02,.2,1) 1;}" +
  "#" + LISTS_LAUNCHER_ID + " .sp-toolkit-launcher-compass.sp-toolkit-launcher-compass-tape-ccw{animation:sp-toolkit-launcher-tape-ccw .52s cubic-bezier(.45,.02,.55,1) 1;}" +
  "@media (prefers-reduced-motion:reduce){#" +
    LISTS_LAUNCHER_ID +
    "{--sp-toolkit-launcher-hover-dur:.2s;--sp-toolkit-launcher-hover-ease:ease;}}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + "{position:absolute;width:min(92vw,1000px);min-width:360px;height:70vh;min-height:280px;max-width:min(92vw,1000px);max-height:70vh;overflow:hidden;display:flex;flex-direction:column;border-radius:11px;box-shadow:0 8px 32px rgba(0,0,0,.25);transition:background .2s,color .2s,border-color .2s;box-sizing:border-box;}" +
  "#" + LISTS_LAUNCHER_ID + ".sp-toolkit-launcher-panel-up #" + LISTS_LAUNCHER_PANEL_ID + "{bottom:calc(var(--sp-toolkit-launcher-size, 58px) + var(--sp-toolkit-launcher-panel-gap, 12px));top:auto;}" +
  "#" + LISTS_LAUNCHER_ID + ".sp-toolkit-launcher-panel-down #" + LISTS_LAUNCHER_PANEL_ID + "{top:calc(var(--sp-toolkit-launcher-size, 58px) + var(--sp-toolkit-launcher-panel-gap, 12px));bottom:auto;}" +
  "#" + LISTS_LAUNCHER_ID + ".sp-toolkit-launcher-panel-anchor-left #" + LISTS_LAUNCHER_PANEL_ID + "{left:0;right:auto;}" +
  "#" + LISTS_LAUNCHER_ID + ".sp-toolkit-launcher-panel-anchor-right #" + LISTS_LAUNCHER_PANEL_ID + "{right:0;left:auto;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark{background:#181b23!important;color:#fff!important;border:1px solid rgba(255,255,255,.08);color-scheme:dark;isolation:isolate;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light{background:#fff;color:#1C1F4A;border:1px solid #E4E6F5;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-lists-panel-header{display:flex;align-items:center;justify-content:space-between;gap:8px;padding:10px 12px;border-bottom:1px solid rgba(255,255,255,.12);flex-shrink:0;transition:border-color .2s;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-lists-panel-header{border-bottom-color:#E4E6F5;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-header-left{display:flex;flex-direction:column;align-items:flex-start;gap:2px;min-width:0;flex:1;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-header-tenant-line{font-size:11px;opacity:.7;color:inherit;margin-top:0;line-height:1.3;max-width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;text-align:left;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-header-tabs{display:flex;align-items:center;gap:2px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-header-tab{padding:4px 10px;font-size:12px;font-weight:500;border-radius:6px;background:transparent;border:none;color:rgba(255,255,255,.6);cursor:pointer;font-family:inherit;transition:background .15s,color .15s;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-header-tab.sp-toolkit-header-tab-label{color:rgba(255,255,255,.9);cursor:default;pointer-events:none;font-size:14px;padding-left:0;text-align:left;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-header-tab.sp-toolkit-header-tab-label .sp-toolkit-label-site-name{font-weight:700;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-header-tab.sp-toolkit-header-tab-label.active{background:transparent;color:rgba(255,255,255,.9);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-header-tab{color:#5A5F8A;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-header-tab.sp-toolkit-header-tab-label{color:#1C1F4A;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-header-tab.sp-toolkit-header-tab-label.active{background:transparent;color:#1C1F4A;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-header-tab:hover{color:inherit;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-header-tab.sp-toolkit-header-tab-label:hover{color:inherit;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-header-tab.active{background:rgba(55,174,28,.2);color:#86efac;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-header-tab.active{background:#EEF0FE;color:#37ae1c;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-subsites-view{display:none;overflow-y:auto;flex:1;min-height:0;padding:12px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-subsites-view.visible{display:block;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-lists-panel-icons{display:flex;align-items:center;gap:4px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-lists-panel-icons a{display:flex;align-items:center;justify-content:center;width:32px;height:32px;border-radius:6px;color:rgba(255,255,255,.85);transition:background .15s,color .15s;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-lists-panel-icons a:hover{background:rgba(55,174,28,.25);color:#86efac;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-lists-panel-icons a{color:#5A5F8A;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-lists-panel-icons a:hover{background:#F5F6FF;color:#37ae1c;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-lists-panel-icons a svg{width:18px;height:18px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-nav-tabs{display:flex;flex-wrap:wrap;align-items:center;gap:4px;margin-top:6px;width:100%;max-width:100%;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-panel-title-label{font-size:14px;font-weight:600;line-height:1.35;color:inherit;text-align:left;max-width:100%;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-panes{position:relative;flex:1;min-height:0;display:flex;flex-direction:column;overflow:hidden;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-pane{position:absolute;inset:0;display:none;flex-direction:column;min-height:0;overflow:hidden;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-pane-active,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-pane.active{display:flex;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-scroll{flex:1;min-height:0;overflow-y:auto;padding:10px 12px;font-size:12px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-site-contents-subbar{display:flex;flex-wrap:wrap;gap:6px;padding:6px 12px;border-bottom:1px solid rgba(255,255,255,.12);flex-shrink:0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-site-contents-subbar{border-bottom-color:#E4E6F5;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-sc-sub{padding:4px 10px;font-size:11px;font-weight:600;border-radius:6px;border:1px solid rgba(255,255,255,.15);background:transparent;color:rgba(255,255,255,.75);cursor:pointer;font-family:inherit;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-sc-sub{border-color:#CDD0EE;color:#5A5F8A;background:#f8f9fd;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-sc-sub.sp-toolkit-sc-sub-active{background:rgba(55,174,28,.2);color:#86efac;border-color:rgba(55,174,28,.4);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-sc-sub.sp-toolkit-sc-sub-active{background:#EEF0FE;color:#37ae1c;border-color:#37ae1c;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-site-contents-stack{flex:1;min-height:0;display:flex;flex-direction:column;overflow:hidden;width:100%;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-site-contents-stack .sp-toolkit-contents-wrap{flex:1;min-height:0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-site-contents-stack .sp-toolkit-subsites-view.visible{flex:1;min-height:0;display:flex;flex-direction:column;overflow-y:auto;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-inline{display:inline-flex;align-items:center;gap:6px;font-size:11px;white-space:nowrap;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-inline input[type=checkbox]{width:16px;height:16px;min-width:16px;min-height:16px;margin:0;flex-shrink:0;accent-color:#37ae1c;cursor:pointer;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-muted{opacity:.85;margin:0;line-height:1.4;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-context-count.sp-toolkit-compass-muted,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-columns-count.sp-toolkit-compass-muted{margin:0!important;opacity:1;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-panel-loading{display:block;padding:10px 12px;font-size:12px;line-height:1.45;color:var(--txt-soft);min-height:36px;box-sizing:border-box;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-err{color:#f87171;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-filter-row{margin-bottom:8px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-filter{width:100%;box-sizing:border-box;padding:7px 10px;border-radius:8px;border:1px solid rgba(255,255,255,.15);background:#252a30;color:inherit;font-size:12px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-ql-filter{background:#fff;border-color:#E4E6F5;color:#1C1F4A;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-section{margin-bottom:8px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-section-label{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;opacity:.75;padding:8px 4px 6px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list{list-style:none;padding:0 0 10px;margin:0;display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list li{margin:0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list a{display:flex;flex-direction:column;align-items:center;justify-content:flex-start;gap:6px;padding:12px 8px 10px;min-height:72px;border-radius:8px;font-size:11px;font-weight:500;color:inherit;text-decoration:none;background:rgba(255,255,255,.03);border:1px solid rgba(255,255,255,.12);text-align:center;line-height:1.25;transition:background .12s,color .12s,border-color .12s,box-shadow .12s;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-ql-list a{background:#fff;border-color:#E4E6F5;color:#1C1F4A;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list a:hover{background:rgba(55,174,28,.12);color:#86efac;border-color:rgba(55,174,28,.4);box-shadow:0 2px 8px rgba(55,174,28,.08);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-ql-list a:hover{background:#EEF0FE;color:#37ae1c;border-color:#C9D3FF;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list a::before{content:'';width:24px;height:24px;display:block;background-size:contain;background-position:center;background-repeat:no-repeat;opacity:.9;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list a[data-ql-icon='settings']::before{background-image:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24' fill='none' stroke='%235A5F8A' stroke-width='2'%3E%3Ccircle cx='12' cy='12' r='3'/%3E%3Cpath d='M19.4 15a1.65 1.65 0 0 0 .33 1.82 2 2 0 1 1-2.83 2.83 1.65 1.65 0 0 0-1.82-.33A1.65 1.65 0 0 0 14 21a2 2 0 1 1-4 0 1.65 1.65 0 0 0-1-1.51 1.65 1.65 0 0 0-1.82.33 2 2 0 1 1-2.83-2.83 1.65 1.65 0 0 0 .33-1.82A1.65 1.65 0 0 0 3 14a2 2 0 1 1 0-4 1.65 1.65 0 0 0 1.51-1 1.65 1.65 0 0 0-.33-1.82 2 2 0 1 1 2.83-2.83A1.65 1.65 0 0 0 9 4.51 1.65 1.65 0 0 0 10 3a2 2 0 1 1 4 0 1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33 2 2 0 1 1 2.83 2.83 1.65 1.65 0 0 0-.33 1.82A1.65 1.65 0 0 0 21 10a2 2 0 1 1 0 4 1.65 1.65 0 0 0-1.6 1z'/%3E%3C/svg%3E\");}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list a[data-ql-icon='trash']::before{background-image:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24' fill='none' stroke='%235A5F8A' stroke-width='2'%3E%3Cpolyline points='3 6 5 6 21 6'/%3E%3Cpath d='M19 6l-1 14H6L5 6m3 0V4h8v2'/%3E%3C/svg%3E\");}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list a[data-ql-icon='users']::before{background-image:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24' fill='none' stroke='%235A5F8A' stroke-width='2'%3E%3Cpath d='M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2'/%3E%3Ccircle cx='9' cy='7' r='4'/%3E%3Cpath d='M23 21v-2a4 4 0 0 0-3-3.87'/%3E%3C/svg%3E\");}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list a[data-ql-icon='chart']::before{background-image:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24' fill='none' stroke='%235A5F8A' stroke-width='2'%3E%3Cline x1='18' y1='20' x2='18' y2='10'/%3E%3Cline x1='12' y1='20' x2='12' y2='4'/%3E%3Cline x1='6' y1='20' x2='6' y2='14'/%3E%3C/svg%3E\");}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list a[data-ql-icon='user']::before{background-image:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24' fill='none' stroke='%235A5F8A' stroke-width='2'%3E%3Ccircle cx='12' cy='7' r='4'/%3E%3Cpath d='M5.5 21a6.5 6.5 0 0 1 13 0'/%3E%3C/svg%3E\");}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list a[data-ql-icon='logout']::before{background-image:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24' fill='none' stroke='%235A5F8A' stroke-width='2'%3E%3Cpath d='M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4'/%3E%3Cpolyline points='16 17 21 12 16 7'/%3E%3Cline x1='21' y1='12' x2='9' y2='12'/%3E%3C/svg%3E\");}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list a[data-ql-icon='monitor']::before{background-image:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24' fill='none' stroke='%235A5F8A' stroke-width='2'%3E%3Crect x='2' y='3' width='20' height='14' rx='2'/%3E%3Cline x1='8' y1='21' x2='16' y2='21'/%3E%3C/svg%3E\");}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list a[data-ql-icon='list']::before{background-image:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24' fill='none' stroke='%235A5F8A' stroke-width='2'%3E%3Cline x1='8' y1='6' x2='21' y2='6'/%3E%3Cline x1='8' y1='12' x2='21' y2='12'/%3E%3Cline x1='8' y1='18' x2='21' y2='18'/%3E%3Ccircle cx='4' cy='6' r='1'/%3E%3Ccircle cx='4' cy='12' r='1'/%3E%3Ccircle cx='4' cy='18' r='1'/%3E%3C/svg%3E\");}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list a[data-ql-icon='layers']::before{background-image:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24' fill='none' stroke='%235A5F8A' stroke-width='2'%3E%3Cpolygon points='12 2 22 8 12 14 2 8 12 2'/%3E%3Cpolyline points='2 14 12 20 22 14'/%3E%3C/svg%3E\");}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list a[data-ql-icon='shield']::before{background-image:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24' fill='none' stroke='%235A5F8A' stroke-width='2'%3E%3Cpath d='M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z'/%3E%3C/svg%3E\");}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list a[data-ql-icon='adminHome']::before{background-image:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24' fill='none' stroke='%235A5F8A' stroke-width='2'%3E%3Crect x='3' y='3' width='7' height='18' rx='1'/%3E%3Cpath d='M14 10l4-4 4 4v11h-8z'/%3E%3C/svg%3E\");}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-context-tools{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:6px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-context-filter{flex:1;min-width:120px;padding:6px 8px;border-radius:6px;border:1px solid rgba(255,255,255,.15);background:#252a30;color:inherit;font-size:12px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-compass-context-filter{background:#fff;border-color:#E4E6F5;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-ctx-list{display:flex;flex-direction:column;gap:0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-ctx-row{display:block;padding:7px 0;border-bottom:1px solid rgba(255,255,255,.08);font-size:11px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-compass-ctx-row{border-bottom-color:#E4E6F5;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-ctx-k{display:block;font-size:11px;font-weight:700;opacity:.85;margin-bottom:3px;font-family:'Courier New',monospace;word-break:break-word;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-ctx-v{word-break:break-word;opacity:.96;max-height:6.5em;overflow:auto;font-family:'Courier New',monospace;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-colhead,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-colrow{display:grid;grid-template-columns:1.25fr 1fr 1fr .8fr;gap:10px;font-size:11px;padding:7px 4px;align-items:start;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-colhead{font-weight:700;border-bottom:1px solid rgba(255,255,255,.15);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-compass-colhead{border-bottom-color:#E4E6F5;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-colrow{border-bottom:1px solid rgba(255,255,255,.06);padding:6px 0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-compass-colrow{border-bottom-color:#EEF0FE;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-colrow .col-name-link{color:inherit;text-decoration:underline;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-colrow .col-name-link:hover{color:#86efac;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-compass-colrow .col-name-link:hover{color:#37ae1c;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-colrow .sp-toolkit-cc-name,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-colrow .col-internal{min-width:0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-colrow .col-internal{font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;word-break:break-all;overflow-wrap:anywhere;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-columns-filter{width:100%;max-width:220px;margin:6px 0;padding:6px 8px;border-radius:6px;border:1px solid rgba(255,255,255,.15);background:#252a30;color:inherit;font-size:12px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-compass-columns-filter{background:#fff;border-color:#E4E6F5;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-label{display:block;font-size:12px;font-weight:600;margin:0;opacity:1;color:inherit;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-primary,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-secondary:not(.btn-ghost-icon){padding:6px 12px;font-size:12px;border-radius:6px;cursor:pointer;font-family:inherit;border:1px solid rgba(255,255,255,.2);background:rgba(55,174,28,.25);color:#86efac;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-compass-primary{background:#37ae1c;color:#fff;border-color:#37ae1c;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-secondary{background:transparent;color:inherit;border-color:rgba(255,255,255,.25);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-btnrow{display:flex;flex-wrap:wrap;gap:8px;margin-top:12px;align-items:center;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-reports-main{display:flex;flex-direction:column;gap:12px;min-height:0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-reports-config{display:flex;flex-direction:column;gap:12px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-field{display:flex;flex-direction:column;gap:6px;align-items:stretch;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-field .sp-toolkit-compass-label{margin:0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-matrix-card{border:1px solid rgba(255,255,255,.12);border-radius:8px;padding:10px 12px;display:flex;flex-direction:column;gap:8px;background:rgba(0,0,0,.12);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-compass-matrix-card{border-color:#E4E6F5;background:#fafbff;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-chk-row{display:flex;align-items:center;gap:8px;font-size:12px;line-height:1.4;margin:0;cursor:pointer;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-chk-row>span{flex:1;min-width:0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-chk-row input[type=checkbox]{width:16px;height:16px;min-width:16px;min-height:16px;margin:0;flex-shrink:0;accent-color:#37ae1c;cursor:pointer;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-matrix-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px 16px;align-items:start;margin-top:0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-matrix-field{display:flex;flex-direction:column;gap:6px;min-width:0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-matrix-field-label{font-size:12px;font-weight:600;line-height:1.35;margin:0;opacity:1;color:var(--pp-txt-mid);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-matrix-field input,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-matrix-field select{width:100%;max-width:100%;box-sizing:border-box;margin:0;padding:0 12px;border-radius:10px;font-size:12px;height:36px;min-height:36px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-matrix-sites-block{display:flex;flex-direction:column;gap:10px;margin-top:4px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-matrix-sites-block > .sp-toolkit-compass-label{margin:0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-matrix-sites-toolbar{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin:0 0 8px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-matrix-sites-toolbar .sp-toolkit-compass-secondary{flex:0 0 auto;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-matrix-sites-list,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-report-lists-list{max-height:min(42vh,360px);min-height:180px;overflow:auto;border:1px solid rgba(255,255,255,.12);border-radius:8px;padding:6px 8px;font-size:12px;line-height:1.4;margin:0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-matrix-sites-list,#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-report-lists-list{border-color:#E4E6F5;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-matrix-site-item{margin:0 0 6px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-export-progress-console{display:none;flex-direction:column;min-height:0;border:1px solid rgba(255,255,255,.15);border-radius:8px;overflow:hidden;background:rgba(0,0,0,.18);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-export-progress-console.sp-toolkit-export-console-visible{display:flex;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-export-progress-console.sp-toolkit-export-console-idle{flex-shrink:0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-export-progress-console.sp-toolkit-export-console-idle .sp-toolkit-export-progress-bar-wrap{display:none;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-export-progress-console.sp-toolkit-export-console-idle .sp-toolkit-export-progress-log{display:none;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-export-progress-console.sp-toolkit-export-console-idle .sp-toolkit-export-progress-head{padding-bottom:8px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-export-progress-headwrap{flex-shrink:0;background:inherit;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-export-progress-console{border-color:#E4E6F5;background:#f8faff;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-export-progress-bar-wrap{height:8px;background:rgba(255,255,255,.08);position:relative;overflow:hidden;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-export-progress-bar-wrap.sp-toolkit-export-bar-active::after{content:'';position:absolute;inset:0;background:linear-gradient(90deg,transparent,rgba(55,174,28,.28),transparent);animation:sp-export-bar-wave 1.8s ease-in-out infinite;pointer-events:none;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-export-progress-bar{height:100%;width:0;background:linear-gradient(90deg,var(--brand-mid),var(--brand));transition:width .25s ease;position:relative;z-index:1;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-export-progress-bar{background:var(--brand);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-export-progress-head{display:flex;justify-content:space-between;gap:10px;align-items:center;padding:8px 10px 4px;font-size:12px;font-weight:600;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-export-progress-status{display:flex;align-items:center;gap:8px;min-width:0;flex:1;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-export-progress-spinner{width:14px;height:14px;border:2px solid rgba(55,174,28,.2);border-top-color:var(--brand-accent);border-radius:50%;flex-shrink:0;display:none;animation:sp-export-spin .75s linear infinite;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-export-progress-console.sp-toolkit-export-console-active .sp-toolkit-export-progress-spinner{display:block;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-export-progress-msg{min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-export-progress-console.sp-toolkit-export-console-active .sp-toolkit-export-progress-msg{color:var(--brand-accent);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-export-progress-meta{display:flex;align-items:center;gap:10px;flex-shrink:0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-export-cancel{display:none;padding:2px 8px;font-size:10px;line-height:1.4;min-height:0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-export-progress-console.sp-toolkit-export-console-active .sp-toolkit-export-cancel{display:inline-flex;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-export-progress-elapsed{font-size:11px;font-weight:600;color:rgba(255,255,255,.55);font-variant-numeric:tabular-nums;min-width:3.5em;text-align:right;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-export-progress-elapsed{color:#64748b;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-export-progress-pct{font-size:11px;font-weight:700;color:var(--brand-accent);white-space:nowrap;}" +
  "@keyframes sp-export-spin{to{transform:rotate(360deg);}}" +
  "@keyframes sp-export-bar-wave{0%{transform:translateX(-100%);}100%{transform:translateX(100%);}}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-export-progress-log{flex:1;min-height:0;max-height:180px;overflow-y:auto;overflow-x:hidden;padding:0 10px 10px;font-family:Consolas,ui-monospace,monospace;font-size:11px;line-height:1.4;color:rgba(255,255,255,.72);white-space:pre-wrap;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-export-progress-log{color:#64748b;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-reports-status.sp-toolkit-reports-status-hidden{display:none!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-reports-main.sp-toolkit-reports-export-active .sp-toolkit-compass-reports-config{display:none!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-reports-main.sp-toolkit-reports-export-active{display:flex;flex-direction:column;flex:1;min-height:240px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-reports-main.sp-toolkit-reports-export-active .sp-toolkit-export-progress-console{flex:1;display:flex!important;flex-direction:column;min-height:0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-reports-main.sp-toolkit-reports-export-active .sp-toolkit-export-progress-log{flex:1;max-height:none;min-height:0;overflow-y:auto;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-scroll.sp-toolkit-reports-scroll-locked{overflow:hidden!important;display:flex;flex-direction:column;min-height:0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-reports-scroll-locked .sp-toolkit-compass-reports-host{flex:1;min-height:0;overflow:hidden;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-reports-scroll-locked .sp-toolkit-compass-reports-main{flex:1;min-height:0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-reports-host select,.sp-toolkit-compass-refinable-type,.sp-toolkit-compass-refinable-num{width:100%;max-width:100%;padding:0 12px;border-radius:10px;font-size:12px;margin:0;display:block;height:36px;min-height:36px;box-sizing:border-box;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-reports-status{margin:0 0 8px;font-size:12px;min-height:1.2em;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-reports-status.ok{opacity:.9;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-picker,.sp-toolkit-compass-settings{margin-top:0;padding-top:0;border-top:none;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-picker-hd{display:flex;align-items:center;gap:10px;margin-bottom:8px;font-weight:600;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-pick-item{margin:4px 0;font-size:12px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-refinable-out{font-size:12px;line-height:1.45;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-refinable-out ul{margin:6px 0 0 18px;padding:0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ref-type{opacity:.75;font-size:11px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-lists-panel-body{display:flex;flex-direction:column;flex:1;min-height:0;overflow:hidden;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-wrap{overflow:auto;flex:1;min-height:0;width:100%;box-sizing:border-box;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table{width:100%;table-layout:fixed;border-collapse:collapse;font-size:12px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table th[data-column=\"name\"]{width:42%;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table th[data-column=\"type\"]{width:22%;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table th[data-column=\"items\"]{width:10%;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table th[data-column=\"modified\"]{width:16%;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table th[data-column=\"settings\"]{width:10%;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-contents-table th{color:rgba(255,255,255,.92)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-contents-table td{color:rgba(255,255,255,.92)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-contents-table .sp-toolkit-row-name a{color:rgba(255,255,255,.95)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-contents-table .sp-toolkit-row-icon{color:rgba(255,255,255,.88)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-contents-table .sp-toolkit-row-settings a{color:rgba(255,255,255,.85)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-contents-table .sp-toolkit-row-view-manager{color:rgba(255,255,255,.85)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-header-tab-label," +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-header-tab-label .sp-toolkit-label-site-name{color:rgba(255,255,255,.92)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-header-tenant-line{color:rgba(255,255,255,.75)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table th{text-align:left;padding:8px 10px;font-weight:600;border-bottom:1px solid rgba(255,255,255,.12);color:inherit;white-space:nowrap;cursor:pointer;user-select:text;-webkit-user-select:text;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table th:first-child," + "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table td:first-child{padding-left:12px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table th:hover{opacity:.9;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table th .sp-toolkit-sort-icon{opacity:.6;margin-left:4px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-contents-table th{border-bottom-color:#E4E6F5;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-site-contents-name-filter{display:flex;align-items:center;gap:6px;max-width:none;width:100%;min-width:0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-site-contents-name-filter .sp-toolkit-sc-filter-icon{width:16px;height:16px;min-width:16px;flex-shrink:0;opacity:.88;color:rgba(255,255,255,.88);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-site-contents-name-filter .sp-toolkit-sc-filter-icon{color:#5A5F8A;opacity:.85;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-site-contents-name-filter .sp-toolkit-sc-filter-icon path{fill:currentColor;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-site-contents-name-filter input{flex:1;min-width:0;width:auto!important;max-width:none!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-filter-row th{padding:4px 8px;vertical-align:middle;cursor:default;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-filter-row th:hover{opacity:1;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-filter-row input," +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-filter-row select{width:100%;max-width:none;padding:6px 8px;font-size:12px;border-radius:6px;border:1px solid rgba(255,255,255,.15);background:#252a30;color:rgba(255,255,255,.95);font-family:inherit;cursor:pointer;appearance:auto;-webkit-appearance:menulist;-moz-appearance:menulist;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-filter-row input{background:#252a30;color:rgba(255,255,255,.95);cursor:text;appearance:none;-webkit-appearance:none;-moz-appearance:none;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-filter-row select:hover{border-color:rgba(255,255,255,.25);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-filter-row select:focus{outline:none;border-color:rgba(55,174,28,.5);box-shadow:0 0 0 2px rgba(55,174,28,.15);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-contents-table .sp-toolkit-filter-row input," +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-contents-table .sp-toolkit-filter-row select{background:#f5f6ff;border-color:#E4E6F5;color:#1C1F4A;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-contents-table .sp-toolkit-filter-row input{background:#fff;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-contents-table .sp-toolkit-filter-row select:hover{border-color:#CDD0EE;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-contents-table .sp-toolkit-filter-row select:focus{border-color:#37ae1c;box-shadow:0 0 0 2px rgba(55,174,28,.2);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-contents-table thead tr.sp-toolkit-filter-row th{position:sticky;top:0;z-index:6;background:#181b23;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-contents-table thead tr.sp-toolkit-filter-row th{position:sticky;top:0;z-index:6;background:#fff;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-contents-table thead tr.sp-toolkit-column-header-row th{position:sticky;top:2.625rem;z-index:5;background:#181b23;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-contents-table thead tr.sp-toolkit-column-header-row th{position:sticky;top:2.625rem;z-index:5;background:#fff;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table td{padding:8px 10px;border-bottom:1px solid rgba(255,255,255,.06);vertical-align:middle;white-space:nowrap;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-contents-table td{border-bottom-color:#E4E6F5;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-row-name{display:flex;align-items:center;gap:8px;min-width:0;overflow:hidden;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-row-icon{flex-shrink:0;width:20px;height:20px;color:inherit;opacity:.85;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-row-name a{color:inherit;text-decoration:none;white-space:nowrap;min-width:0;overflow:hidden;text-overflow:ellipsis;flex:1;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-contents-table .sp-toolkit-row-name a:hover{color:#86efac;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-contents-table .sp-toolkit-row-name a:hover{color:#37ae1c;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table tr.sp-toolkit-current .sp-toolkit-row-name a{font-weight:600;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table tr.sp-toolkit-current td:first-child{border-left:3px solid #37ae1c;padding-left:9px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table td.sp-toolkit-type-cell{display:flex;align-items:center;color:inherit;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-type-pill{display:inline-flex;align-items:center;justify-content:center;line-height:1;font-size:11px;font-weight:700;padding:3px 8px;border-radius:999px;white-space:nowrap;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-type-pill{background:rgba(55,174,28,.25);color:#86efac;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-type-pill{background:#EEF0FE;color:#5A5F8A;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table td:nth-child(3),#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table td:nth-child(4){text-align:center;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-row-settings{display:flex;align-items:center;justify-content:center;gap:2px;min-width:56px;text-align:center;vertical-align:middle;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-row-settings a{display:inline-flex;align-items:center;justify-content:center;color:inherit;opacity:.7;padding:4px;border-radius:4px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-row-settings a:hover{opacity:1;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-row-settings svg{width:18px;height:18px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-row-view-manager{display:inline-flex;align-items:center;justify-content:center;color:inherit;opacity:.7;padding:4px;margin:0;border:none;background:transparent;cursor:pointer;border-radius:4px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-row-view-manager:hover{opacity:1;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-row-view-manager svg{width:18px;height:18px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table th[data-column=\"settings\"]{cursor:default;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-row-name a:hover{text-decoration:underline;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-lists-loading,.sp-toolkit-lists-error{padding:14px;font-size:13px;color:#9ca3bf;transition:color .2s;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-lists-loading,.sp-toolkit-lists-panel-light .sp-toolkit-lists-error{color:#5A5F8A;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-open{animation:sp-toolkit-panel-fade-in .2s ease-out forwards;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-panel-rolling-up{animation:sp-toolkit-panel-fade-out .12s ease-out forwards;}" +
  "@keyframes sp-toolkit-panel-fade-in{0%{opacity:0;}100%{opacity:1;}}" +
  "@keyframes sp-toolkit-panel-fade-out{0%{opacity:1;}100%{opacity:0;}}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-lists-loading{display:flex;align-items:center;gap:10px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-lists-loading-dot{width:6px;height:6px;border-radius:50%;background:rgba(55,174,28,.8);animation:sp-toolkit-load-bounce 1s ease-in-out infinite;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-lists-loading-dot:nth-child(2){animation-delay:.12s;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-lists-loading-dot:nth-child(3){animation-delay:.24s;}" +
  "@keyframes sp-toolkit-load-bounce{0%,80%,100%{transform:scale(0.6);opacity:.6;}40%{transform:scale(1);opacity:1;}}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-pane-content-enter{animation:none!important;}" +
  "@keyframes sp-toolkit-pane-fade-y{0%{opacity:0;transform:translateY(-5px);}100%{opacity:1;transform:translateY(0);}}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-flow-list > *{opacity:1!important;transform:none!important;animation:none!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-flow-list .sp-toolkit-ql-list li," +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-flow-list .sp-toolkit-compass-ctx-row," +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-flow-list .sp-toolkit-compass-colrow," +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-flow-list .refinable-mappings-list li{opacity:1!important;transform:none!important;animation:none!important;}" +
  "@keyframes sp-toolkit-flow-down{0%{opacity:0;transform:translateY(-6px);}100%{opacity:1;transform:translateY(0);}}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table.sp-toolkit-lists-loaded tbody tr{animation:sp-toolkit-item-in .3s ease-out backwards;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table.sp-toolkit-lists-loaded tbody tr:nth-child(1){animation-delay:0.02s;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table.sp-toolkit-lists-loaded tbody tr:nth-child(2){animation-delay:0.04s;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table.sp-toolkit-lists-loaded tbody tr:nth-child(3){animation-delay:0.06s;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table.sp-toolkit-lists-loaded tbody tr:nth-child(4){animation-delay:0.08s;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table.sp-toolkit-lists-loaded tbody tr:nth-child(5){animation-delay:0.1s;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table.sp-toolkit-lists-loaded tbody tr:nth-child(n+6){animation-delay:0.12s;}" +
  "@keyframes sp-toolkit-item-in{0%{opacity:0;transform:translateX(-6px);}100%{opacity:1;transform:translateX(0);}}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + "{backdrop-filter:blur(6px);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-lists-panel-header{padding:12px 14px 10px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-panel-title-label{font-size:13px;font-weight:700;letter-spacing:.01em;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-header-tenant-line{font-size:10px;opacity:.68;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-nav-tabs{gap:6px;margin-top:8px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-pane-tab{padding:5px 10px;border-radius:999px;font-weight:600;font-size:11px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-compass-pane-tab{background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.08);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-compass-pane-tab.active{background:rgba(55,174,28,.2);border-color:rgba(55,174,28,.45);color:#b8ffc8;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-compass-pane-tab{background:#f8f9ff;border:1px solid #e3e6f8;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-compass-pane-tab.active{background:#e8f7e4;border-color:#7cca66;color:#2d8d1b;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-scroll{padding:12px 14px 14px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-context-filter,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-columns-filter,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-filter,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-reports-host select,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-matrix-field input,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-matrix-field select,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-refinable-type,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-refinable-num{height:32px;border-radius:8px;font-size:12px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-primary,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-secondary:not(.btn-ghost-icon){height:32px;padding:0 12px;border-radius:8px;font-weight:600;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-compass-secondary{background:rgba(255,255,255,.04);border-color:rgba(255,255,255,.16);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list a[data-ql-icon]::before{opacity:.95;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list a[data-ql-icon='wrench']::before{background-image:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24' fill='none' stroke='%235A5F8A' stroke-width='2'%3E%3Cpath d='M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.7-3.8a6 6 0 0 1-7.9 7.9l-6.9 6.9a2.1 2.1 0 1 1-3-3l6.9-6.9a6 6 0 0 1 7.9-7.9z'/%3E%3C/svg%3E\");}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list a[data-ql-icon='key']::before{background-image:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24' fill='none' stroke='%235A5F8A' stroke-width='2'%3E%3Ccircle cx='8.5' cy='15.5' r='4.5'/%3E%3Cpath d='M13 11l8-8m-2 0h2v2'/%3E%3C/svg%3E\");}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list a[data-ql-icon='search']::before{background-image:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24' fill='none' stroke='%235A5F8A' stroke-width='2'%3E%3Ccircle cx='11' cy='11' r='7'/%3E%3Cline x1='20' y1='20' x2='16.6' y2='16.6'/%3E%3C/svg%3E\");}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list a[data-ql-icon='teams']::before{background-image:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24' fill='none' stroke='%235A5F8A' stroke-width='2'%3E%3Crect x='3' y='4' width='12' height='16' rx='2'/%3E%3Cpath d='M15 9h6v6h-6z'/%3E%3C/svg%3E\");}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list a[data-ql-icon='package']::before{background-image:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24' fill='none' stroke='%235A5F8A' stroke-width='2'%3E%3Cpath d='M3 8l9-5 9 5-9 5-9-5z'/%3E%3Cpath d='M3 8v8l9 5 9-5V8'/%3E%3C/svg%3E\");}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list a[data-ql-icon='archive']::before{background-image:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24' fill='none' stroke='%235A5F8A' stroke-width='2'%3E%3Crect x='3' y='4' width='18' height='5'/%3E%3Cpath d='M5 9h14v11H5z'/%3E%3C/svg%3E\");}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list a[data-ql-icon='tag']::before{background-image:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24' fill='none' stroke='%235A5F8A' stroke-width='2'%3E%3Cpath d='M20 13l-7 7a2 2 0 0 1-3 0l-7-7V3h10l7 10z'/%3E%3Ccircle cx='7.5' cy='7.5' r='1.2'/%3E%3C/svg%3E\");}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list a[data-ql-icon]::before{transition:transform .12s;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list a:hover::before{transform:scale(1.08);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-ctx-row{padding:8px 0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-ctx-v{font-size:11px;line-height:1.35;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-colhead{position:sticky;top:0;z-index:2;background:inherit;padding-top:6px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-colrow{font-size:11px;line-height:1.3;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-columns-list{border:1px solid rgba(255,255,255,.08);border-radius:10px;overflow:auto;padding:0 8px 6px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-compass-columns-list{border-color:#e6e9fb;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-reports-host,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-refinable-wrap,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-views-host{background:rgba(255,255,255,.02);border:1px solid rgba(255,255,255,.08);border-radius:10px;padding:10px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-compass-reports-host,#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-compass-refinable-wrap,#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-compass-views-host{background:#fff;border-color:#e6e9fb;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-reports-status{padding:6px 8px;border-radius:8px;background:rgba(255,255,255,.04);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-compass-reports-status{background:#f5f7ff;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + "{--pp-bg:#101722;--pp-surface:#182231;--pp-surface-2:#1d2838;--pp-border:rgba(255,255,255,.12);--pp-txt:#f3f6fb;--pp-txt-mid:#c5cfdb;--pp-txt-soft:#9aa8ba;--pp-brand:#37ae1c;--pp-brand-soft:rgba(55,174,28,.35);--pp-brand-faint:rgba(55,174,28,.12);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light{--pp-bg:#ffffff;--pp-surface:#ffffff;--pp-surface-2:#f7f9ff;--pp-border:#e3e8f4;--pp-txt:#1f2a3a;--pp-txt-mid:#5a6780;--pp-txt-soft:#7e8ca3;--pp-brand:#37ae1c;--pp-brand-soft:#a9df9d;--pp-brand-faint:#ecf9e8;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + "{background:var(--pp-bg)!important;color:var(--pp-txt)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-lists-panel-header{background:var(--pp-surface);border-bottom:1px solid var(--pp-border)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-header-tenant-line{color:var(--pp-txt-soft)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-pane-tab{background:transparent!important;border:0!important;color:var(--pp-txt-mid)!important;padding:4px 10px!important;border-radius:6px!important;font-size:12px!important;font-weight:600!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-pane-tab:hover{background:transparent!important;color:var(--pp-txt)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-pane-tab.active{background:var(--pp-brand-faint)!important;color:var(--pp-brand)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-scroll{background:var(--pp-bg);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-section-label{color:var(--pp-txt-soft)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list a{background:var(--pp-surface)!important;border:1.5px solid var(--pp-border)!important;color:var(--pp-txt)!important;min-height:74px!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-list a:hover{background:var(--pp-brand-faint)!important;border-color:var(--pp-brand-soft)!important;color:var(--pp-brand)!important;box-shadow:0 2px 8px rgba(55,174,28,.08)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-context-filter,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-columns-filter,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-filter,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-reports-host select,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-matrix-field input,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-matrix-field select,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-refinable-type,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-refinable-num{background:var(--pp-surface-2)!important;border:1.5px solid var(--pp-border)!important;color:var(--pp-txt)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-context-filter:focus,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-columns-filter:focus,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-ql-filter:focus,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-reports-host select:focus,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-refinable-type:focus,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-refinable-num:focus{outline:none;border-color:var(--pp-brand-soft)!important;box-shadow:0 0 0 2px rgba(55,174,28,.15)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-ctx-row{border-bottom:1px solid var(--pp-border)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-ctx-k{color:var(--pp-txt-mid)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-ctx-v{color:var(--pp-txt)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-columns-count,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-context-count{color:var(--pp-txt-soft)!important;font-size:11px!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-columns-list{background:var(--pp-surface)!important;border:1px solid var(--pp-border)!important;border-radius:8px!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-column-creator-host:empty{display:none!important;margin:0!important;padding:0!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-column-creator-host:not(:empty){margin:0 0 8px!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator{margin:0;padding:12px 14px;border:1.5px solid var(--pp-border);border-radius:10px;background:var(--pp-surface);display:flex;flex-direction:column;gap:12px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator[hidden]{display:none!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-label{display:block;font-size:12px;font-weight:600;color:var(--pp-txt-mid);margin:0;line-height:1.35;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-check{display:flex;align-items:center;gap:8px;font-size:12px;color:var(--pp-txt);margin:0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-create-fields,#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-site-pick{display:flex;flex-direction:column;gap:12px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-type-options{display:flex;flex-direction:column;gap:8px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-create-fields[hidden],#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-site-pick[hidden]{display:none!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-preview{padding:8px 10px;border-radius:8px;background:var(--pp-surface-2);border:1px solid var(--pp-border);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-preview-label{display:block;font-size:10px;font-weight:700;color:var(--pp-txt-soft);text-transform:uppercase;letter-spacing:.04em;margin-bottom:4px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-internal{display:block;font-size:12px;font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;color:var(--pp-txt);word-break:break-all;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-crawled{margin-top:4px;font-size:11px;color:var(--pp-txt-soft);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-notes{font-size:11px;color:var(--pp-txt-mid);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-note{margin:0 0 4px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-actions{display:flex;align-items:center;gap:10px;flex-wrap:wrap;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-submit{padding:7px 12px;border-radius:8px;border:none;background:var(--pp-brand);color:#fff;font-size:12px;font-weight:600;cursor:pointer;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-submit:disabled{opacity:.55;cursor:not-allowed;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-status{font-size:11px;color:var(--pp-txt-soft);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-status.err{color:#f87171;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-status.ok{color:#4ade80;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-placement,#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-type,#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-group,#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-title-input,#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-description,#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-choices,#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-site-filter,#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-type-options select,#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-type-options input{width:100%;box-sizing:border-box;padding:8px 11px;border-radius:8px;border:1.5px solid var(--pp-border);background:var(--pp-surface-2);color:var(--pp-txt);font-size:12px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-description,#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-choices{min-height:72px;resize:vertical;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-site-list{display:flex;flex-direction:column;gap:6px;max-height:220px;overflow:auto;border:1px solid var(--pp-border);border-radius:8px;background:var(--pp-surface-2);padding:6px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-site-row{display:grid;grid-template-columns:minmax(0,1.2fr) minmax(0,1fr) auto auto;gap:8px;align-items:center;padding:6px 8px;border-radius:6px;background:var(--pp-surface);border:1px solid var(--pp-border);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-site-title{font-size:12px;font-weight:600;color:var(--pp-txt);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-site-internal{font-size:11px;font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;color:var(--pp-txt-mid);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-site-type{font-size:10px;color:var(--pp-txt-soft);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-site-add{padding:5px 10px;border-radius:6px;border:1px solid var(--pp-brand-soft);background:var(--pp-brand-faint);color:var(--pp-brand);font-size:11px;font-weight:600;cursor:pointer;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-site-add:disabled{opacity:.55;cursor:not-allowed;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-site-empty{padding:10px 8px;font-size:11px;color:var(--pp-txt-soft);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " #btnOpenColumnCreator.column-creator-open,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-btn-open-column-creator.column-creator-open{background:var(--pp-brand-faint)!important;border-color:var(--pp-brand-soft)!important;color:var(--pp-brand)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .column-creator-load-err{padding:10px 12px;font-size:12px;color:var(--pp-txt-soft);border:1px dashed var(--pp-border);border-radius:8px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-colhead{font-size:11px!important;color:var(--pp-txt-mid)!important;border-bottom:1px solid var(--pp-border)!important;text-transform:none!important;letter-spacing:0!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-colrow{border-bottom:1px solid var(--pp-border)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-colrow:last-child{border-bottom:none!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-colrow .col-name-link{color:var(--pp-txt)!important;text-decoration:none!important;font-weight:500;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-colrow .col-name-link:hover{color:var(--pp-brand)!important;text-decoration:underline!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-primary{background:var(--pp-brand)!important;color:#fff!important;border:1px solid var(--pp-brand)!important;box-shadow:0 3px 10px rgba(55,174,28,.24)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-primary:hover{filter:brightness(.95);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-secondary{background:var(--pp-surface-2)!important;color:var(--pp-txt-mid)!important;border:1.5px solid var(--pp-border)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-secondary:hover{background:var(--pp-brand-faint)!important;color:var(--pp-brand)!important;border-color:var(--pp-brand-soft)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-reports-host,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-refinable-wrap,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-views-host{background:transparent!important;border:0!important;padding:0!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-reports-status{background:var(--pp-surface-2)!important;border:1px solid var(--pp-border);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-pane[data-compass-pane='refinableProps'] .sp-toolkit-compass-scroll{max-width:none!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-pane[data-compass-pane='refinableProps'] .sp-toolkit-compass-refinable-type,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-pane[data-compass-pane='refinableProps'] .sp-toolkit-compass-refinable-num{width:100%!important;max-width:100%!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-site-contents-subbar{background:var(--pp-surface);border-bottom:1px solid var(--pp-border)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-sc-sub{background:var(--pp-surface-2)!important;border:1px solid var(--pp-border)!important;color:var(--pp-txt-mid)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-sc-sub.sp-toolkit-sc-sub-active{background:var(--pp-brand-faint)!important;border-color:var(--pp-brand-soft)!important;color:var(--pp-brand)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table thead tr.sp-toolkit-filter-row th,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table thead tr.sp-toolkit-column-header-row th{background:var(--pp-surface)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table th,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table td{border-bottom:1px solid var(--pp-border)!important;color:var(--pp-txt)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-filter-row input,#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-filter-row select{background:var(--pp-surface-2)!important;border:1px solid var(--pp-border)!important;color:var(--pp-txt)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " #quicklinksContent{display:block!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-pane[data-compass-pane='quicklinks'] .sp-toolkit-compass-scroll{padding:0!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-pane[data-compass-pane='quicklinks'] .sp-toolkit-compass-quicklinks-host{padding:10px 12px 12px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " #contextPanel,#" + LISTS_LAUNCHER_PANEL_ID + " #searchSchemaPanel{display:flex!important;flex-direction:column!important;min-height:0!important;gap:0!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " #refinablePropsPanel{padding:0!important;display:flex;flex-direction:column;gap:12px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .context-row .name{font-size:11px;font-weight:700;color:var(--pp-txt-mid)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .context-value-row{font-family:'Courier New',monospace;font-size:11px;line-height:1.35;color:var(--pp-txt)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " #searchSchemaList .search-schema-row{padding:8px 12px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " #btnRefreshContext,#" + LISTS_LAUNCHER_PANEL_ID + " #btnOpenColumnCreator,#" + LISTS_LAUNCHER_PANEL_ID + " #btnRefinableConfigure{font-size:12px;font-weight:600;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " #refinableMappingsOut{min-height:44px;padding:6px 0;color:var(--pp-txt-soft)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " #refinableMappingsOut .sp-toolkit-panel-loading{padding:8px 0;min-height:32px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-nav-tabs{border-bottom:none;padding-bottom:0;margin-bottom:0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-pane-tab{border-bottom:2px solid transparent!important;border-radius:0!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-pane-tab.active{background:transparent!important;border-bottom-color:var(--pp-brand)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-compass-pane-tab.active:hover{background:transparent!important;color:var(--pp-brand)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " #refinablePropsPanel{padding:0!important;display:flex;flex-direction:column;gap:12px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " #refinablePropsPanel .sp-toolkit-compass-label{text-transform:none!important;letter-spacing:0!important;font-size:12px!important;color:var(--pp-txt-mid)!important;margin:0!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " #refinablePropsPanel .sp-toolkit-compass-muted{color:var(--pp-txt-soft)!important;font-size:12px;line-height:1.45;margin:2px 0 0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " #refinableTypeSelect,#" + LISTS_LAUNCHER_PANEL_ID + " #refinableNumberSelect{height:36px!important;border-radius:10px!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " #refinableTypeSelect:hover,#" + LISTS_LAUNCHER_PANEL_ID + " #refinableNumberSelect:hover{border-color:var(--pp-brand-soft)!important;color:var(--pp-brand)!important;background:var(--pp-brand-faint)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " #btnRefinableConfigure{height:32px!important;min-width:150px!important;padding:0 12px!important;border-radius:8px!important;display:inline-flex;align-items:center;justify-content:center;gap:8px;font-weight:600!important;align-self:flex-start;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " #btnRefinableConfigure .sp-toolkit-btn-gear{font-size:13px;line-height:1;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " #btnRefinableConfigure:hover{transform:translateY(-1px);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .refinable-mappings-header{font-size:11px;font-weight:700;color:var(--pp-txt-soft);text-transform:uppercase;letter-spacing:.05em;display:block;margin-bottom:6px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .refinable-mappings-list{list-style:none;padding:0;margin:0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .refinable-mappings-list li{padding:4px 0;border-bottom:1px solid var(--pp-border);font-family:monospace;font-size:11px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .refinable-mappings-list li:last-child{border-bottom:none;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-compass-pane-tab:hover{background:transparent!important;color:#d2dced!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-compass-pane-tab:hover{background:transparent!important;color:#3c4c66!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-compass-pane-tab.active:hover{color:#37ae1c!important;background:transparent!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-compass-pane-tab.active:hover{color:#37ae1c!important;background:transparent!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-ql-list a:hover{background:rgba(123,137,245,.14)!important;border-color:rgba(123,137,245,.38)!important;color:#dbe6ff!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-ql-list a:hover{background:#eef2ff!important;border-color:#cfd7ff!important;color:#3d4fd6!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-compass-context-filter:hover,#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-compass-columns-filter:hover,#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-ql-filter:hover,#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-compass-reports-host select:hover,#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-compass-refinable-type:hover,#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-compass-refinable-num:hover{border-color:rgba(123,137,245,.45)!important;background:rgba(123,137,245,.12)!important;color:#dbe6ff!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-compass-context-filter:hover,#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-compass-columns-filter:hover,#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-ql-filter:hover,#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-compass-reports-host select:hover,#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-compass-refinable-type:hover,#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-compass-refinable-num:hover{border-color:#c7d1ff!important;background:#f1f4ff!important;color:#3d4fd6!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-compass-secondary:hover{background:rgba(123,137,245,.16)!important;color:#dbe6ff!important;border-color:rgba(123,137,245,.42)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-compass-secondary:hover{background:#eef2ff!important;color:#3d4fd6!important;border-color:#cfd7ff!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-compass-primary:hover{background:#2aa642!important;border-color:#2aa642!important;box-shadow:0 6px 18px rgba(42,166,66,.28)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-compass-primary:hover{background:#2faa45!important;border-color:#2faa45!important;box-shadow:0 6px 18px rgba(42,166,66,.24)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-compass-ctx-row:hover{background:rgba(123,137,245,.08)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-compass-ctx-row:hover{background:#f6f8ff!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-compass-colrow:hover{background:rgba(123,137,245,.08)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-compass-colrow:hover{background:#f6f8ff!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-compass-colrow .col-name-link:hover{color:#bfcfff!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-compass-colrow .col-name-link:hover{color:#3d4fd6!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-sc-sub:hover{background:rgba(123,137,245,.14)!important;border-color:rgba(123,137,245,.4)!important;color:#dbe6ff!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-sc-sub:hover{background:#eef2ff!important;border-color:#cfd7ff!important;color:#3d4fd6!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-contents-table tbody tr:hover td{background:rgba(123,137,245,.06)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-contents-table tbody tr:hover td{background:#f7f9ff!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-contents-table .sp-toolkit-row-settings a:hover,#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-contents-table .sp-toolkit-row-view-manager:hover{background:rgba(123,137,245,.16)!important;color:#dbe6ff!important;opacity:1!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-contents-table .sp-toolkit-row-settings a:hover,#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-contents-table .sp-toolkit-row-view-manager:hover{background:#eef2ff!important;color:#3d4fd6!important;opacity:1!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-subsites-list a:hover{background:rgba(123,137,245,.2)!important;color:#dbe6ff!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-subsites-list a:hover{background:#e8eeff!important;color:#3d4fd6!important;}";

const LISTS_LAUNCHER_PARITY_CSS = `
#${LISTS_LAUNCHER_PANEL_ID}{
  --bg:#0f172a;
  --surface:#111827;
  --surface-2:#1f2937;
  --border:#243246;
  --txt:#e6edf6;
  --txt-mid:#b4c0d2;
  --txt-soft:#8ea0b8;
  --brand:#37ae1c;
  --brand-dark:#2f9518;
  --brand-mid:#5fd247;
  --brand-accent:#86efac;
  --brand-faint:rgba(55,174,28,.14);
  --r-md:8px;
  --r-lg:12px;
}
#${LISTS_LAUNCHER_PANEL_ID}.sp-toolkit-lists-panel-light{
  --bg:#ffffff;
  --surface:#ffffff;
  --surface-2:#f8faff;
  --border:#e2e8f4;
  --txt:#1f2a3a;
  --txt-mid:#52627c;
  --txt-soft:#7f8ea5;
  --brand:#37ae1c;
  --brand-dark:#2f9518;
  --brand-mid:#5fd247;
  --brand-accent:#37ae1c;
  --brand-faint:rgba(55,174,28,.14);
}
#${LISTS_LAUNCHER_PANEL_ID}{
  background:var(--bg)!important;
  color:var(--txt)!important;
  border:1px solid var(--border)!important;
  border-radius:12px!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-lists-panel-header{
  padding:10px 14px 8px!important;
  background:var(--surface)!important;
  border-bottom:1px solid var(--border)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-header-left{
  align-self:stretch!important;
  width:100%!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-header-title-row{
  display:flex!important;
  align-items:center!important;
  justify-content:flex-start!important;
  gap:8px!important;
  width:100%!important;
  align-self:stretch!important;
  box-sizing:border-box!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-header-title-row .sp-toolkit-panel-title-label{
  flex:1!important;
  min-width:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-header-actions{
  display:flex!important;
  align-items:center!important;
  gap:8px!important;
  flex:0 0 auto!important;
  margin-left:auto!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-toolbar-icon-btn.sp-toolkit-compass-universal-search-btn,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-toolbar-icon-btn.sp-toolkit-compass-settings-btn,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-toolbar-icon-btn.sp-toolkit-compass-close-btn{
  box-sizing:border-box!important;
  width:auto!important;
  height:28px!important;
  min-width:28px!important;
  padding:0 9px!important;
  margin:0!important;
  border-radius:10px!important;
  border:1.5px solid var(--border)!important;
  background:var(--surface-2)!important;
  color:var(--txt-mid)!important;
  display:inline-flex!important;
  align-items:center!important;
  justify-content:center!important;
  cursor:pointer!important;
  transition:background .18s ease, border-color .18s ease, color .18s ease, box-shadow .18s ease!important;
  box-shadow:0 1px 0 rgba(255,255,255,.04) inset!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-toolbar-icon-btn.sp-toolkit-compass-close-btn{
  font-size:16px!important;
  font-weight:500!important;
  line-height:1!important;
  padding:0!important;
  width:28px!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-toolbar-icon-btn.sp-toolkit-compass-universal-search-btn:hover,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-toolbar-icon-btn.sp-toolkit-compass-settings-btn:hover,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-toolbar-icon-btn.sp-toolkit-compass-close-btn:hover{
  background:var(--brand-faint)!important;
  color:var(--brand)!important;
  border-color:var(--brand-mid)!important;
  box-shadow:0 0 0 1px rgba(55,174,28,.12)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-toolbar-icon-btn.sp-toolkit-compass-universal-search-btn:focus-visible,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-toolbar-icon-btn.sp-toolkit-compass-settings-btn:focus-visible,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-toolbar-icon-btn.sp-toolkit-compass-close-btn:focus-visible{
  outline:2px solid var(--brand-mid)!important;
  outline-offset:2px!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-header-search-svg{
  width:17px!important;
  height:17px!important;
  min-width:17px!important;
  display:block!important;
  flex-shrink:0!important;
  opacity:.95!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-header-title-row .dark-toggle{
  display:flex!important;
  align-items:center!important;
  flex:0 0 52px!important;
  width:52px!important;
  min-width:52px!important;
  max-width:52px!important;
  margin:0!important;
  padding:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .dark-toggle-btn{
  width:52px!important;
  min-width:52px!important;
  max-width:52px!important;
  height:24px;
  border:none;
  border-radius:12px;
  cursor:pointer;
  position:relative;
  flex-shrink:0;
  overflow:hidden;
  padding:0!important;
  margin:0!important;
  box-sizing:border-box!important;
  background:linear-gradient(90deg,#f5e6c8 0%,#e8d4a8 35%,#3d2a5c 65%,#1a0a2e 100%);
  box-shadow:inset 0 1px 2px rgba(0,0,0,.2),0 0 0 1px rgba(255,255,255,.15);
  transition:box-shadow .2s ease;
}
#${LISTS_LAUNCHER_PANEL_ID} .dark-toggle-btn:hover{
  box-shadow:inset 0 1px 2px rgba(0,0,0,.25),0 0 0 1px rgba(255,255,255,.25);
}
#${LISTS_LAUNCHER_PANEL_ID} .dark-toggle-btn::before{
  content:'';
  position:absolute;
  left:5px;
  top:50%;
  transform:translateY(-50%);
  width:14px;
  height:14px;
  background:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23b8860b' stroke-width='1.5' stroke-linecap='round' stroke-linejoin='round'%3E%3Ccircle cx='12' cy='10' r='4'/%3E%3Cpath d='M12 2v2M12 18v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M18 12h2M4.93 19.07l1.41-1.41M17.66 6.34l1.41-1.41'/%3E%3Cpath d='M5 14h14'/%3E%3C/svg%3E") no-repeat center / contain;
  opacity:.95;
  pointer-events:none;
}
#${LISTS_LAUNCHER_PANEL_ID} .dark-toggle-btn::after{
  content:'';
  position:absolute;
  right:5px;
  top:50%;
  transform:translateY(-50%);
  width:14px;
  height:14px;
  background:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23c9b8e8' stroke-width='1.5' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z'/%3E%3C/svg%3E") no-repeat center / contain;
  opacity:.9;
  pointer-events:none;
}
#${LISTS_LAUNCHER_PANEL_ID} .dark-toggle-btn .dark-toggle-knob{
  position:absolute;
  top:2px;
  left:2px;
  width:20px;
  height:20px;
  border-radius:50%;
  background:radial-gradient(circle at 30% 30%,#fff 0%,#f0ebe0 50%,#e0d8c8 100%);
  box-shadow:0 1px 3px rgba(0,0,0,.35),0 0 0 1px rgba(255,255,255,.5);
  transition:transform .25s cubic-bezier(0.4,0,0.2,1);
  pointer-events:none;
  display:flex;
  align-items:center;
  justify-content:center;
}
#${LISTS_LAUNCHER_PANEL_ID} .dark-toggle-btn .dark-toggle-knob::before{
  content:'';
  width:11px;
  height:11px;
  background:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23c4952a' stroke-width='1.8' stroke-linecap='round' stroke-linejoin='round'%3E%3Ccircle cx='12' cy='9' r='3'/%3E%3Cpath d='M12 1v1.5M12 19.5V21M4.22 4.22l1.06 1.06M18.72 18.72l1.06 1.06M1 12h1.5M19.5 12H21M4.22 19.78l1.06-1.06M18.72 5.28l1.06-1.06'/%3E%3Cpath d='M4 14h16'/%3E%3C/svg%3E") no-repeat center / contain;
}
#${LISTS_LAUNCHER_PANEL_ID}.sp-toolkit-lists-panel-dark .dark-toggle-btn .dark-toggle-knob::before{
  content:'';
  width:10px;
  height:10px;
  background:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23e8e0f0' stroke-width='1.5' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z'/%3E%3C/svg%3E") no-repeat center / contain;
}
#${LISTS_LAUNCHER_PANEL_ID}.sp-toolkit-lists-panel-dark .dark-toggle-btn::after{
  opacity:0;
}
#${LISTS_LAUNCHER_PANEL_ID}.sp-toolkit-lists-panel-dark .dark-toggle-btn .dark-toggle-knob{
  transform:translateX(28px);
  background:radial-gradient(circle at 30% 30%,#c9b8e8 0%,#5c4d7a 50%,#2d2345 100%);
  box-shadow:0 1px 3px rgba(0,0,0,.5),0 0 0 1px rgba(255,255,255,.15);
}
#${LISTS_LAUNCHER_PANEL_ID} .dark-toggle-btn:focus-visible{
  outline:2px solid white;
  outline-offset:2px;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-panel-title-label{
  font-size:15px!important;
  font-weight:700!important;
  line-height:1.3!important;
  color:var(--txt)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-title-sep{
  font-weight:400!important;
  color:var(--txt-soft)!important;
  opacity:.75!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-title-suffix{
  font-weight:400!important;
  color:var(--txt-mid)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-header-tenant-line{
  font-size:11px!important;
  color:var(--txt-soft)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-nav-tabs{
  display:flex!important;
  gap:0!important;
  margin-top:8px!important;
  border-bottom:none!important;
  overflow-x:auto!important;
  overflow-y:visible!important;
  position:relative!important;
  padding-bottom:3px!important;
  scroll-behavior:smooth!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .tab-bar{
  -ms-overflow-style:none;
  scrollbar-width:none;
}
#${LISTS_LAUNCHER_PANEL_ID} .tab-bar::-webkit-scrollbar{display:none;}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane-tab{
  flex:1 1 auto!important;
  padding:8px 10px!important;
  background:transparent!important;
  border:none!important;
  border-bottom:2px solid transparent!important;
  border-radius:0!important;
  color:var(--txt-mid)!important;
  font-size:12px!important;
  font-weight:600!important;
  white-space:nowrap!important;
  transition:color .12s ease, background .12s ease!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane-tab:hover{
  color:var(--txt)!important;
  background:transparent!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane-tab.active{
  color:var(--brand)!important;
  border-bottom-color:transparent!important;
  background:transparent!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane-tab.active:hover{
  color:var(--brand)!important;
  background:transparent!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane-tab:disabled{
  opacity:.45!important;
  cursor:not-allowed!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane-tab:not(:disabled){
  transition:color .2s ease!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-tab-indicator{
  position:absolute!important;
  left:0;
  bottom:0;
  height:2.5px;
  width:28px;
  border-radius:2px 2px 0 0;
  background:linear-gradient(90deg,var(--brand-mid),var(--brand));
  transition:left .22s cubic-bezier(.4,0,.2,1), width .22s cubic-bezier(.4,0,.2,1), opacity .16s ease;
  opacity:0;
  pointer-events:none;
  z-index:4;
  box-shadow:0 0 10px rgba(55,174,28,.35);
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-tab-indicator.sp-toolkit-tab-indicator-instant{
  transition:none!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-scroll{
  background:var(--bg)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-panes{
  position:relative!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane.tab-content{
  display:none!important;
  position:absolute!important;
  inset:0!important;
  flex-direction:column!important;
  min-height:0!important;
  overflow:hidden!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane.tab-content.active,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane.tab-content.sp-toolkit-compass-pane-active{
  display:flex!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-label,
#${LISTS_LAUNCHER_PANEL_ID} .form-label{
  display:block!important;
  margin:0!important;
  font-size:12px!important;
  font-weight:600!important;
  letter-spacing:0!important;
  text-transform:none!important;
  color:var(--txt-mid)!important;
  line-height:1.35!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-field,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-matrix-field{
  gap:var(--compass-field-gap)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-matrix-field-label{
  font-size:12px!important;
  font-weight:600!important;
  color:var(--txt-mid)!important;
  margin:0!important;
  line-height:1.35!important;
}
#${LISTS_LAUNCHER_PANEL_ID} input:not([type="checkbox"]):not([type="radio"]):not([type="range"]),
#${LISTS_LAUNCHER_PANEL_ID} select{
  width:100%!important;
  min-height:36px!important;
  padding:8px 12px!important;
  box-sizing:border-box!important;
  line-height:1.35!important;
  border-radius:10px!important;
  border:1.5px solid var(--border)!important;
  background:var(--surface-2)!important;
  color:var(--txt)!important;
  font-size:12px!important;
  transition:border-color .3s cubic-bezier(.22,1,.36,1), background .3s cubic-bezier(.22,1,.36,1), box-shadow .3s cubic-bezier(.22,1,.36,1), transform .25s cubic-bezier(.22,1,.36,1)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} select{
  appearance:none!important;
  -webkit-appearance:none!important;
  -moz-appearance:none!important;
  cursor:pointer!important;
  padding:8px 34px 8px 12px!important;
  background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 20 20' fill='%238ea0b8'%3E%3Cpath fill-rule='evenodd' d='M5.23 7.21a.75.75 0 011.06.02L10 11.164l3.71-3.934a.75.75 0 111.08 1.04l-4.25 4.5a.75.75 0 01-1.08 0l-4.25-4.5a.75.75 0 01.02-1.06z' clip-rule='evenodd'/%3E%3C/svg%3E")!important;
  background-repeat:no-repeat!important;
  background-size:12px 12px!important;
  background-position:right 11px center!important;
}
#${LISTS_LAUNCHER_PANEL_ID} select::-ms-expand{
  display:none;
}
#${LISTS_LAUNCHER_PANEL_ID}.sp-toolkit-lists-panel-dark select{
  color-scheme:dark;
}
#${LISTS_LAUNCHER_PANEL_ID}.sp-toolkit-lists-panel-light select{
  color-scheme:light;
}
#${LISTS_LAUNCHER_PANEL_ID}.sp-toolkit-lists-panel-dark select option,
#${LISTS_LAUNCHER_PANEL_ID}.sp-toolkit-lists-panel-dark select optgroup{
  background-color:#121b2b;
  color:#e6edf6;
}
#${LISTS_LAUNCHER_PANEL_ID}.sp-toolkit-lists-panel-light select option,
#${LISTS_LAUNCHER_PANEL_ID}.sp-toolkit-lists-panel-light select optgroup{
  background-color:#ffffff;
  color:#1f2a3a;
}
#${LISTS_LAUNCHER_PANEL_ID} select:disabled{
  opacity:.6!important;
  cursor:not-allowed!important;
}
#${LISTS_LAUNCHER_PANEL_ID} input:not([type="checkbox"]):not([type="radio"]):hover,
#${LISTS_LAUNCHER_PANEL_ID} select:hover{
  border-color:var(--brand-mid)!important;
  background:var(--brand-faint)!important;
  box-shadow:0 4px 14px rgba(55,174,28,.08)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} input:not([type="checkbox"]):not([type="radio"]):focus,
#${LISTS_LAUNCHER_PANEL_ID} select:focus{
  outline:none!important;
  border-color:var(--brand-mid)!important;
  box-shadow:0 0 0 2px rgba(55,174,28,.2)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} input[type="checkbox"],
#${LISTS_LAUNCHER_PANEL_ID} input[type="radio"]{
  width:auto!important;
  min-height:auto!important;
  max-height:none!important;
  border:none!important;
  background:transparent!important;
  box-shadow:none!important;
  border-radius:4px!important;
  accent-color:var(--brand)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-context-byvalue{
  width:16px!important;
  height:16px!important;
  min-width:16px!important;
  min-height:16px!important;
  flex-shrink:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-picker{
  display:none!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-picker.sp-toolkit-compass-picker-open{
  display:flex!important;
  flex-direction:column!important;
  gap:8px!important;
  min-height:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-settings{
  display:none!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-matrix-settings-section{
  margin-top:4px;
  padding-top:12px;
  border-top:1px solid var(--border);
  display:flex;
  flex-direction:column;
  gap:8px;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-matrix-settings-section > .sp-toolkit-compass-label{
  margin:0 0 2px!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-chk-row{
  align-items:center!important;
  line-height:1.4!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-chk-row>span{
  flex:1 1 auto!important;
  min-width:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-chk-row input[type=checkbox]{
  width:16px!important;
  height:16px!important;
  min-width:16px!important;
  min-height:16px!important;
  max-width:16px!important;
  max-height:16px!important;
  margin:0!important;
  flex-shrink:0!important;
  cursor:pointer!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-settings.sp-toolkit-compass-settings-open{
  display:flex!important;
  flex-direction:column!important;
  gap:12px!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-picker-hd{
  flex-shrink:0!important;
  margin-bottom:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-picker-status{
  flex-shrink:0!important;
  font-size:12px!important;
  min-height:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-column-list{
  flex:0 1 auto!important;
  min-height:88px!important;
  max-height:min(46vh,380px)!important;
  overflow-y:auto!important;
  overflow-x:hidden!important;
  padding:4px!important;
  margin:0!important;
  border:1px solid var(--border)!important;
  border-radius:10px!important;
  background:var(--surface)!important;
  -webkit-overflow-scrolling:touch!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pick-item{
  display:flex!important;
  align-items:flex-start!important;
  gap:10px!important;
  padding:5px 8px!important;
  margin:0!important;
  border-radius:8px!important;
  cursor:pointer!important;
  transition:background .22s cubic-bezier(.22,1,.36,1)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pick-item:hover{
  background:var(--brand-faint)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pick-item input[type="checkbox"]{
  width:16px!important;
  height:16px!important;
  min-width:16px!important;
  min-height:16px!important;
  max-width:16px!important;
  max-height:16px!important;
  margin:3px 0 0!important;
  flex-shrink:0!important;
  cursor:pointer!important;
  accent-color:var(--brand)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pick-item label{
  flex:1 1 auto!important;
  min-width:0!important;
  margin:0!important;
  padding:0!important;
  font-size:12px!important;
  font-weight:500!important;
  line-height:1.35!important;
  color:var(--txt)!important;
  cursor:pointer!important;
  display:block!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pick-item .sp-toolkit-pick-internal{
  display:block!important;
  margin-top:3px!important;
  font-size:10px!important;
  font-weight:500!important;
  font-family:'Segoe UI',system-ui,sans-serif!important;
  color:var(--txt-soft)!important;
  word-break:break-word!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pick-item .sp-toolkit-pick-req{
  color:var(--brand)!important;
  font-weight:700!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-reports-host{
  display:flex!important;
  flex-direction:column!important;
  gap:12px!important;
  min-height:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-primary,
#${LISTS_LAUNCHER_PANEL_ID} .btn-configure{
  display:inline-flex!important;
  align-items:center!important;
  justify-content:center!important;
  gap:8px!important;
  min-height:38px!important;
  padding:9px 16px!important;
  border-radius:10px!important;
  border:1px solid var(--brand)!important;
  background:var(--brand)!important;
  color:#fff!important;
  font-size:13px!important;
  font-weight:700!important;
  box-shadow:0 3px 10px rgba(55,174,28,.3)!important;
  transition:transform .32s cubic-bezier(.22,1,.36,1), box-shadow .32s cubic-bezier(.22,1,.36,1), background .28s ease, border-color .28s ease, filter .22s ease!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-primary:hover,
#${LISTS_LAUNCHER_PANEL_ID} .btn-configure:hover{
  background:var(--brand-dark)!important;
  border-color:var(--brand-dark)!important;
  transform:translateY(-2px)!important;
  box-shadow:0 10px 26px rgba(55,174,28,.4)!important;
  filter:saturate(1.05)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-primary:active,
#${LISTS_LAUNCHER_PANEL_ID} .btn-configure:active{
  transform:translateY(0)!important;
  box-shadow:0 3px 12px rgba(55,174,28,.28)!important;
  transition-duration:.12s!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-secondary:not(.btn-ghost-icon){
  min-height:36px!important;
  padding:8px 12px!important;
  border-radius:8px!important;
  border:1.5px solid var(--border)!important;
  background:var(--surface-2)!important;
  color:var(--txt-mid)!important;
  font-size:12px!important;
  font-weight:600!important;
  transition:transform .3s cubic-bezier(.22,1,.36,1), box-shadow .3s cubic-bezier(.22,1,.36,1), border-color .28s ease, background .28s ease, color .28s ease!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-secondary:not(.btn-ghost-icon):hover{
  border-color:var(--brand-mid)!important;
  color:var(--brand)!important;
  background:var(--brand-faint)!important;
  transform:translateY(-2px)!important;
  box-shadow:0 8px 20px rgba(55,174,28,.1)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-secondary:not(.btn-ghost-icon):active{
  transform:translateY(0)!important;
  box-shadow:0 2px 8px rgba(0,0,0,.06)!important;
  transition-duration:.12s!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-ql-section-label{
  font-size:10px!important;
  font-weight:700!important;
  letter-spacing:.07em!important;
  text-transform:uppercase!important;
  color:var(--txt-soft)!important;
  padding:10px 2px 6px!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #quicklinksCurrentSite,
#${LISTS_LAUNCHER_PANEL_ID} #quicklinksCurrentUser,
#${LISTS_LAUNCHER_PANEL_ID} #quicklinksModes,
#${LISTS_LAUNCHER_PANEL_ID} #quicklinksTenant{
  display:grid!important;
  grid-template-columns:repeat(2,minmax(0,1fr))!important;
  gap:8px!important;
  padding:0 0 10px!important;
  margin:0!important;
  list-style:none!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-ql-list a{
  min-height:72px!important;
  padding:10px 8px!important;
  border-radius:10px!important;
  border:1.5px solid var(--border)!important;
  background:var(--surface)!important;
  color:var(--txt)!important;
  font-size:11px!important;
  font-weight:600!important;
  text-decoration:none!important;
  transition:transform .34s cubic-bezier(.22,1,.36,1), box-shadow .34s cubic-bezier(.22,1,.36,1), border-color .3s ease, background .3s ease, color .28s ease!important;
  will-change:transform;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-ql-list a:hover{
  border-color:var(--brand-mid)!important;
  background:var(--brand-faint)!important;
  color:var(--brand)!important;
  transform:translateY(-4px) scale(1.02)!important;
  box-shadow:0 12px 28px rgba(0,0,0,.14)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-ql-list a:active{
  transform:translateY(-1px) scale(1.01)!important;
  transition-duration:.14s!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaList,#${LISTS_LAUNCHER_PANEL_ID} #contextList{
  border-left:1px solid var(--border)!important;
  border-right:1px solid var(--border)!important;
  border-bottom:1px solid var(--border)!important;
  border-radius:0 0 10px 10px!important;
  overflow:auto!important;
  background:var(--surface)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #contextPanel,
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel{
  display:flex!important;
  flex-direction:column!important;
  min-height:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #contextList,
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaList{
  flex:1 1 auto!important;
  min-height:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaList .sp-toolkit-panel-loading,#${LISTS_LAUNCHER_PANEL_ID} #contextList .sp-toolkit-panel-loading{
  margin:0;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaList .search-schema-row{
  border-bottom:1px solid var(--border)!important;
  padding:var(--compass-list-row-pad-y) var(--compass-list-header-pad-x)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #contextList .search-schema-row,
#${LISTS_LAUNCHER_PANEL_ID} #contextList .compass-page-props-row{
  border-bottom:1px solid var(--border)!important;
  padding:var(--compass-list-row-pad-y) var(--compass-list-header-pad-x)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaList .search-schema-row:last-child,#${LISTS_LAUNCHER_PANEL_ID} #contextList .search-schema-row:last-child{
  border-bottom:none!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-colhead{
  position:sticky!important;
  top:0!important;
  background:var(--surface)!important;
  z-index:2!important;
  color:var(--txt-mid)!important;
  font-weight:700!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .toolbar{
  display:flex;
  align-items:center;
  gap:var(--compass-list-toolbar-gap);
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .toolbar #searchSchemaFilter{
  flex:1 1 auto;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .toolbar .btn-ghost-icon{
  width:36px;
  min-width:36px;
  height:36px;
  padding:0!important;
  border-radius:10px!important;
  display:inline-flex;
  align-items:center;
  justify-content:center;
  border:1.5px solid transparent!important;
  background:var(--surface-2)!important;
  transition:transform .28s cubic-bezier(.22,1,.36,1), box-shadow .28s cubic-bezier(.22,1,.36,1), border-color .25s ease, background .25s ease, color .25s ease!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .toolbar .btn-ghost-icon:hover{
  transform:translateY(-2px) scale(1.04)!important;
  border-color:var(--brand-mid)!important;
  background:var(--brand-faint)!important;
  color:var(--brand)!important;
  box-shadow:0 6px 16px rgba(55,174,28,.12)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .toolbar .btn-ghost-icon:active{
  transform:translateY(0) scale(1)!important;
  transition-duration:.12s!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .toolbar .btn-ghost-icon svg{
  width:16px;
  height:16px;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .schema-col-header{
  display:grid;
  grid-template-columns:2fr 2fr 2fr 1fr;
  gap:8px;
  padding:var(--compass-list-header-pad-y) var(--compass-list-header-pad-x);
  border:1px solid var(--border);
  border-radius:10px 10px 0 0;
  border-bottom:none;
  background:var(--surface);
  margin:0;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .schema-col-header span{
  font-size:10px;
  font-weight:700;
  color:var(--txt-soft);
  text-transform:uppercase;
  letter-spacing:.05em;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel #searchSchemaList .search-schema-row{
  display:grid;
  grid-template-columns:2fr 2fr 2fr 1fr;
  gap:8px;
  align-items:start;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .col-name-link{
  color:var(--brand)!important;
  font-weight:600!important;
  text-decoration:none!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .col-name-link:hover{
  text-decoration:underline!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .sp-toolkit-cc-name{
  min-width:0;
  word-break:break-word;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .col-internal{
  font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;
  font-size:11px;
  color:var(--txt-mid);
  min-width:0;
  word-break:break-all;
  overflow-wrap:anywhere;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .col-crawled{
  font-size:11px;
  color:var(--txt-soft);
  white-space:nowrap;
  overflow:hidden;
  text-overflow:ellipsis;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .col-type{
  justify-self:start;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .col-type-pill{
  display:inline-flex;
  align-items:center;
  padding:2px 7px;
  border-radius:999px;
  font-size:10px;
  font-weight:700;
  background:var(--brand-faint);
  color:var(--brand);
  white-space:nowrap;
}
#${LISTS_LAUNCHER_PANEL_ID} #contextPanel .toolbar{
  display:flex;
  align-items:center;
  gap:var(--compass-list-toolbar-gap);
}
#${LISTS_LAUNCHER_PANEL_ID} #contextPanel .toolbar #contextFilterInput{
  flex:1 1 auto;
}
#${LISTS_LAUNCHER_PANEL_ID} #contextPanel .toolbar .btn-ghost-icon{
  width:36px;
  min-width:36px;
  height:36px;
  padding:0!important;
  border-radius:10px!important;
  display:inline-flex;
  align-items:center;
  justify-content:center;
  border:1.5px solid transparent!important;
  background:var(--surface-2)!important;
  transition:transform .28s cubic-bezier(.22,1,.36,1), box-shadow .28s cubic-bezier(.22,1,.36,1), border-color .25s ease, background .25s ease, color .25s ease!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #contextPanel .toolbar .btn-ghost-icon:hover{
  transform:translateY(-2px) scale(1.04)!important;
  border-color:var(--brand-mid)!important;
  background:var(--brand-faint)!important;
  color:var(--brand)!important;
  box-shadow:0 6px 16px rgba(55,174,28,.12)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #contextPanel .toolbar .btn-ghost-icon:active{
  transform:translateY(0) scale(1)!important;
  transition-duration:.12s!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #contextPanel .toolbar .btn-ghost-icon svg{
  width:16px;
  height:16px;
}
#${LISTS_LAUNCHER_PANEL_ID} #contextPanel .toolbar .sp-toolkit-compass-secondary.btn-ghost-icon,
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .toolbar .sp-toolkit-compass-secondary.btn-ghost-icon{
  display:inline-flex!important;
  align-items:center!important;
  justify-content:center!important;
  line-height:0!important;
  padding:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #contextPanel .toolbar .sp-toolkit-compass-secondary.btn-ghost-icon svg,
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .toolbar .sp-toolkit-compass-secondary.btn-ghost-icon svg{
  display:block!important;
  flex-shrink:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #contextPanel .compass-context-schema-head{
  display:grid;
  grid-template-columns:minmax(0,1fr) minmax(0,2fr);
  gap:8px;
  padding:var(--compass-list-header-pad-y) var(--compass-list-header-pad-x);
  border:1px solid var(--border);
  border-radius:10px 10px 0 0;
  border-bottom:none;
  background:var(--surface);
  margin:0;
}
#${LISTS_LAUNCHER_PANEL_ID} #contextPanel .compass-context-schema-head span{
  font-size:10px;
  font-weight:700;
  color:var(--txt-soft);
  text-transform:uppercase;
  letter-spacing:.05em;
}
#${LISTS_LAUNCHER_PANEL_ID} #contextList .compass-page-props-row{
  display:grid;
  grid-template-columns:minmax(0,1fr) minmax(0,2fr);
  gap:8px;
  align-items:center;
}
#${LISTS_LAUNCHER_PANEL_ID} #contextList .compass-page-props-name{
  font-size:11px;
  font-weight:600!important;
  color:var(--brand)!important;
  word-break:break-word;
}
#${LISTS_LAUNCHER_PANEL_ID} #contextList .compass-page-props-value-wrap{
  display:flex;
  align-items:flex-start;
  gap:6px;
  min-width:0;
}
#${LISTS_LAUNCHER_PANEL_ID} #contextList .compass-page-props-value{
  font-family:monospace;
  font-size:11px;
  color:var(--txt);
  flex:1;
  min-width:0;
  word-break:break-word;
  max-height:6.5em;
  overflow:auto;
  border-radius:4px;
  outline:1px solid transparent;
  transition:outline-color .15s ease, background-color .15s ease;
}
#${LISTS_LAUNCHER_PANEL_ID} #contextList .compass-page-props-value-wrap:hover .compass-page-props-value{
  outline-color:rgba(55,174,28,.25);
  background:rgba(55,174,28,.06);
}
#${LISTS_LAUNCHER_PANEL_ID} #contextList .compass-page-props-value-copied{
  outline-color:var(--brand-mid)!important;
  background:var(--brand-faint)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #contextList .compass-page-props-copy-btn{
  width:20px;
  min-width:20px;
  height:20px;
  border:1px solid transparent;
  border-radius:6px;
  background:transparent;
  color:var(--txt-soft);
  display:inline-flex;
  align-items:center;
  justify-content:center;
  cursor:pointer;
  padding:0;
}
#${LISTS_LAUNCHER_PANEL_ID} #contextList .compass-page-props-copy-btn svg{
  width:13px;
  height:13px;
}
#${LISTS_LAUNCHER_PANEL_ID} #contextList .compass-page-props-copy-btn:hover{
  color:var(--brand);
  border-color:var(--border);
  background:var(--brand-faint);
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-colrow{
  transition:background-color .3s cubic-bezier(.22,1,.36,1)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-colrow:hover{
  background:var(--surface-2)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #reportsPanel,
#${LISTS_LAUNCHER_PANEL_ID} #refinablePropsPanel,
#${LISTS_LAUNCHER_PANEL_ID} #viewManagerPanel{
  padding:14px!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #refinableMappingsOut{
  margin-top:8px!important;
  max-height:180px!important;
  overflow-y:auto!important;
  color:var(--txt-mid)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .refinable-mappings-header{
  font-size:11px!important;
  font-weight:700!important;
  color:var(--txt-soft)!important;
  text-transform:uppercase!important;
  letter-spacing:.05em!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .refinable-mappings-list li{
  padding:4px 0!important;
  border-bottom:1px solid var(--border)!important;
  font-family:monospace!important;
  font-size:11px!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-site-contents-subbar{
  background:var(--surface)!important;
  border-bottom:1px solid var(--border)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-sc-sub{
  border:1.5px solid var(--border)!important;
  background:var(--surface-2)!important;
  color:var(--txt-mid)!important;
  transition:transform .28s cubic-bezier(.22,1,.36,1), box-shadow .28s cubic-bezier(.22,1,.36,1), border-color .26s ease, background .26s ease, color .26s ease!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-sc-sub:hover{
  border-color:var(--brand-mid)!important;
  color:var(--brand)!important;
  background:var(--brand-faint)!important;
  transform:translateY(-2px)!important;
  box-shadow:0 6px 16px rgba(55,174,28,.1)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-sc-sub:active{
  transform:translateY(0)!important;
  transition-duration:.12s!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-sc-sub.sp-toolkit-sc-sub-active{
  border-color:var(--brand-mid)!important;
  color:var(--brand)!important;
  background:var(--brand-faint)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-contents-table th,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-contents-table td{
  border-bottom:1px solid var(--border)!important;
  color:var(--txt)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-contents-table tbody td{
  transition:background-color .32s cubic-bezier(.22,1,.36,1)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-contents-table tbody tr:hover td{
  background:var(--surface-2)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane[data-compass-pane='quicklinks'] .sp-toolkit-compass-scroll{
  padding:12px 14px 14px!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane[data-compass-pane='quicklinks'] .sp-toolkit-compass-quicklinks-host{
  padding:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane[data-compass-pane='quicklinks'] .sp-toolkit-ql-filter{
  min-height:34px!important;
  border-radius:10px!important;
  padding:8px 12px!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane[data-compass-pane='quicklinks'] .sp-toolkit-ql-section-label{
  padding:12px 2px 8px!important;
  font-size:12px!important;
  letter-spacing:.06em!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane[data-compass-pane='quicklinks'] .sp-toolkit-ql-list{
  gap:10px!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane[data-compass-pane='quicklinks'] .sp-toolkit-ql-list a{
  min-height:74px!important;
  border-radius:12px!important;
  border:1px solid #2a3850!important;
  background:linear-gradient(180deg,#111a28 0%,#0f1723 100%)!important;
  color:#e8eef8!important;
  font-size:13px!important;
  font-weight:700!important;
  box-shadow:inset 0 1px 0 rgba(255,255,255,.03), 0 2px 8px rgba(0,0,0,.2)!important;
  justify-content:center!important;
  transition:transform .36s cubic-bezier(.22,1,.36,1), box-shadow .36s cubic-bezier(.22,1,.36,1), border-color .3s ease, background .32s ease, color .28s ease, filter .28s ease!important;
  will-change:transform;
}
#${LISTS_LAUNCHER_PANEL_ID}.sp-toolkit-lists-panel-light .sp-toolkit-compass-pane[data-compass-pane='quicklinks'] .sp-toolkit-ql-list a{
  border-color:#dbe3f2!important;
  background:linear-gradient(180deg,#ffffff 0%,#f8faff 100%)!important;
  color:#223149!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane[data-compass-pane='quicklinks'] .sp-toolkit-ql-list a::before{
  width:30px!important;
  height:30px!important;
  opacity:.9!important;
  margin-top:2px!important;
  margin-bottom:2px!important;
  transition:filter .32s ease, transform .32s cubic-bezier(.22,1,.36,1)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane[data-compass-pane='quicklinks'] .sp-toolkit-ql-list a:hover{
  background:linear-gradient(180deg,#10251b 0%,#102017 100%)!important;
  border-color:#2f8c4a!important;
  color:#d8ffe5!important;
  box-shadow:0 0 0 1px rgba(56,212,93,.25), inset 0 0 0 1px rgba(56,212,93,.12), 0 14px 32px rgba(0,0,0,.28)!important;
  transform:translateY(-5px) scale(1.025)!important;
}
#${LISTS_LAUNCHER_PANEL_ID}.sp-toolkit-lists-panel-light .sp-toolkit-compass-pane[data-compass-pane='quicklinks'] .sp-toolkit-ql-list a:hover{
  background:linear-gradient(180deg,#edf9f0 0%,#e6f6eb 100%)!important;
  border-color:#8ac89a!important;
  color:#1f7a36!important;
  box-shadow:0 10px 26px rgba(55,174,28,.14)!important;
  transform:translateY(-4px) scale(1.02)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane[data-compass-pane='quicklinks'] .sp-toolkit-ql-list a:active{
  transform:translateY(-1px) scale(1.01)!important;
  transition-duration:.14s!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane[data-compass-pane='quicklinks'] .sp-toolkit-ql-list a:hover::before{
  filter:brightness(1.15) saturate(1.9) hue-rotate(-25deg);
  transform:scale(1.06)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #quicklinksFilter,
#${LISTS_LAUNCHER_PANEL_ID} #contextFilterInput,
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaFilter,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-ql-filter,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-context-filter,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-columns-filter,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-report-select,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-format-select,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-page-size,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-default-format,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-matrix-max-items,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-matrix-page-size,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-refinable-type,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-refinable-num{
  width:100%!important;
  max-width:100%!important;
  height:36px!important;
  min-height:36px!important;
  max-height:36px!important;
  line-height:1.35!important;
  padding:8px 12px!important;
  box-sizing:border-box!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-report-select,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-format-select,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-page-size,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-default-format,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-matrix-page-size,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-refinable-type,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-refinable-num{
  padding-right:34px!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .toolbar .btn-ghost-icon,
#${LISTS_LAUNCHER_PANEL_ID} #btnRefreshContext,
#${LISTS_LAUNCHER_PANEL_ID} #btnOpenColumnCreator,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-btn-refresh-context,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-btn-open-column-creator{
  height:36px!important;
  min-height:36px!important;
  max-height:36px!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #contextPanel .toolbar,
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .toolbar,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-ql-filter-row{
  display:flex!important;
  gap:var(--compass-list-toolbar-gap)!important;
  min-height:36px!important;
  align-items:center!important;
  margin:0!important;
  padding:0 0 var(--compass-list-toolbar-pad-bottom)!important;
  border-bottom:1px solid var(--border)!important;
  flex-shrink:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #quicklinksFilter,
#${LISTS_LAUNCHER_PANEL_ID} #contextFilterInput,
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaFilter{
  flex:1 1 auto!important;
  margin:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-ql-filter-row::after{
  content:"";
  width:36px;
  min-width:36px;
  height:36px;
  border-radius:10px;
  border:1.5px solid transparent;
  opacity:0;
  pointer-events:none;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane[data-compass-pane='quicklinks'] .sp-toolkit-ql-filter-row::after{
  display:none;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane[data-compass-pane='quicklinks'] .sp-toolkit-ql-filter-row{
  border-bottom:none!important;
  padding:0!important;
  margin:0 0 10px!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaCount,
#${LISTS_LAUNCHER_PANEL_ID} #contextCount,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-columns-count,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-context-count{
  margin:0!important;
  padding:0 0 var(--compass-list-count-pad-bottom)!important;
  min-height:16px!important;
  flex-shrink:0!important;
  display:block!important;
  font-size:11px!important;
  color:var(--txt-soft)!important;
  line-height:1.35!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-column-creator-host:empty{
  display:none!important;
  margin:0!important;
  padding:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-column-creator-host:not(:empty){
  margin:0 0 var(--compass-list-count-pad-bottom)!important;
  padding:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .sp-toolkit-compass-columns-error{
  margin:0 0 var(--compass-list-count-pad-bottom)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .sp-toolkit-compass-columns-error:empty,
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .sp-toolkit-compass-columns-error[style*="display: none"],
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .sp-toolkit-compass-columns-error[style*="display:none"]{
  display:none!important;
  margin:0!important;
  padding:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-reports-status:empty{
  display:none!important;
  padding:0!important;
  border:0!important;
  margin:0!important;
  min-height:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane[data-compass-pane='reports'] .sp-toolkit-compass-scroll{
  padding-bottom:16px!important;
  -webkit-mask-image:none!important;
  mask-image:none!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane[data-compass-pane='reports'] .sp-toolkit-compass-scroll.sp-toolkit-reports-matrix-scroll{
  display:flex!important;
  flex-direction:column!important;
  min-height:0!important;
  overflow:hidden!important;
  padding-bottom:12px!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-matrix-mode.sp-toolkit-compass-reports-host,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-matrix-mode .sp-toolkit-compass-reports-main,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-matrix-mode .sp-toolkit-compass-reports-config,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-matrix-mode .sp-toolkit-compass-report-opts,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-matrix-mode .sp-toolkit-matrix-row,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-matrix-mode .sp-toolkit-matrix-sites-block{
  flex:1!important;
  min-height:0!important;
  display:flex!important;
  flex-direction:column!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-matrix-mode .sp-toolkit-compass-reports-status,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-matrix-mode .sp-toolkit-compass-reports-config > .sp-toolkit-compass-field{
  flex:0 0 auto!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-matrix-mode .sp-toolkit-matrix-sites-toolbar{
  flex:0 0 auto!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-matrix-mode .sp-toolkit-matrix-sites-block > .sp-toolkit-compass-label{
  flex:0 0 auto!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane[data-compass-pane='reports'] .sp-toolkit-compass-scroll.sp-toolkit-reports-lists-scroll{
  display:flex!important;
  flex-direction:column!important;
  min-height:0!important;
  overflow:hidden!important;
  padding-bottom:12px!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-lists-mode.sp-toolkit-compass-reports-host,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-lists-mode .sp-toolkit-compass-reports-main,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-lists-mode .sp-toolkit-compass-reports-config,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-lists-mode .sp-toolkit-compass-report-opts,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-lists-mode .sp-toolkit-report-lists-row,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-lists-mode .sp-toolkit-report-lists-block{
  flex:1!important;
  min-height:0!important;
  display:flex!important;
  flex-direction:column!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-lists-mode .sp-toolkit-compass-reports-status,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-lists-mode .sp-toolkit-compass-reports-config > .sp-toolkit-compass-field:not(.sp-toolkit-report-lists-block),
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-lists-mode .sp-toolkit-report-lists-toolbar,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-lists-mode .sp-toolkit-report-lists-filter{
  flex:0 0 auto!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-reports-main.sp-toolkit-reports-export-tracking .sp-toolkit-compass-reports-config{
  display:none!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-reports-main.sp-toolkit-reports-export-tracking .sp-toolkit-export-progress-console{
  flex:1!important;
  display:flex!important;
  flex-direction:column!important;
  min-height:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-lists-mode .sp-toolkit-compass-reports-config{
  flex:1 1 auto!important;
  min-height:0!important;
  overflow:hidden!important;
  display:flex!important;
  flex-direction:column!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-lists-mode .sp-toolkit-compass-report-opts{
  flex:1 1 auto!important;
  min-height:0!important;
  overflow:hidden!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-lists-mode .sp-toolkit-report-lists-row,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-lists-mode .sp-toolkit-report-lists-block{
  flex:1 1 auto!important;
  min-height:0!important;
  overflow:hidden!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-lists-mode .sp-toolkit-report-lists-list{
  flex:1 1 auto!important;
  min-height:0!important;
  max-height:none!important;
  overflow-y:auto!important;
  overflow-x:hidden!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-reports-footer{
  flex:0 0 auto!important;
  display:flex!important;
  flex-direction:column!important;
  gap:8px!important;
  margin-top:8px!important;
  padding-top:10px!important;
  border-top:1px solid var(--border,rgba(255,255,255,.12))!important;
  background:var(--bg)!important;
  position:relative!important;
  z-index:3!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-bundle-hint{
  margin:0!important;
  font-size:11px!important;
  line-height:1.35!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-lists-mode .sp-toolkit-compass-reports-footer,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-lists-mode .sp-toolkit-compass-reports-footer .sp-toolkit-export-row,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-lists-mode .sp-toolkit-compass-reports-footer .sp-toolkit-compass-btnrow{
  flex:0 0 auto!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-lists-filter{
  width:100%!important;
  max-width:100%!important;
  box-sizing:border-box!important;
  margin:0 0 8px!important;
  padding:0 12px!important;
  height:36px!important;
  min-height:36px!important;
  border-radius:8px!important;
  font-size:12px!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-lists-count{
  margin-left:auto!important;
  font-size:11px!important;
  font-weight:600!important;
  opacity:.85!important;
  white-space:nowrap!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-lists-list.sp-toolkit-report-lists-tree{
  padding:4px 6px 8px!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-tree-node{
  --tree-indent:14px;
  margin:0 0 4px!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-tree-site-heading{
  display:flex!important;
  align-items:flex-start!important;
  gap:4px!important;
  margin:0 0 2px!important;
  padding:6px 6px 6px 2px!important;
  border-radius:8px!important;
  background:var(--surface-2,rgba(255,255,255,.06))!important;
  border:1px solid var(--border,rgba(255,255,255,.1))!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-tree-toggle{
  flex:0 0 auto!important;
  display:inline-flex!important;
  align-items:center!important;
  justify-content:center!important;
  width:22px!important;
  height:28px!important;
  margin:0!important;
  padding:0!important;
  border:none!important;
  border-radius:6px!important;
  background:transparent!important;
  color:inherit!important;
  cursor:pointer!important;
  opacity:.85!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-tree-toggle:hover{
  background:rgba(55,174,28,.12)!important;
  opacity:1!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-tree-chevron{
  display:block!important;
  width:0!important;
  height:0!important;
  border-top:5px solid transparent!important;
  border-bottom:5px solid transparent!important;
  border-left:6px solid currentColor!important;
  transition:transform .16s cubic-bezier(.4,0,.2,1)!important;
  transform:rotate(0deg)!important;
  transform-origin:40% 50%!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-tree-node:not(.sp-toolkit-report-tree-collapsed) > .sp-toolkit-report-tree-site-heading .sp-toolkit-report-tree-chevron{
  transform:rotate(90deg)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-tree-site-label{
  display:flex!important;
  align-items:flex-start!important;
  gap:8px!important;
  min-width:0!important;
  flex:1!important;
  margin:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-tree-site-cb{
  margin-top:6px!important;
  flex-shrink:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-tree-site-text{
  display:flex!important;
  flex-direction:column!important;
  gap:2px!important;
  min-width:0!important;
  flex:1!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-tree-site-meta{
  font-size:10px!important;
  line-height:1.35!important;
  opacity:.72!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-tree-body{
  margin:0 0 6px 0!important;
  padding-left:calc(var(--tree-indent) + 6px)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-tree-node.sp-toolkit-report-tree-collapsed > .sp-toolkit-report-tree-body{
  display:none!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-tree-lists{
  display:flex!important;
  flex-direction:column!important;
  gap:2px!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-lists-site-heading{
  position:sticky!important;
  top:0!important;
  z-index:2!important;
  margin:10px 0 6px!important;
  padding:8px 8px 6px!important;
  border-radius:8px!important;
  background:var(--surface-2,rgba(255,255,255,.06))!important;
  border:1px solid var(--border,rgba(255,255,255,.1))!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-lists-site-heading:first-child{
  margin-top:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-lists-site-title{
  font-size:12px!important;
  font-weight:700!important;
  line-height:1.35!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-lists-site-path{
  margin-top:2px!important;
  font-size:10px!important;
  line-height:1.35!important;
  opacity:.78!important;
  word-break:break-all!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-list-item{
  margin:0 0 4px!important;
  padding:6px 8px!important;
  border-radius:8px!important;
  border:1px solid transparent!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-list-item:hover{
  background:rgba(55,174,28,.08)!important;
  border-color:rgba(55,174,28,.18)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-list-item label{
  align-items:flex-start!important;
  gap:10px!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-list-item input{
  margin-top:3px!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-list-item-main{
  display:flex!important;
  flex-direction:column!important;
  gap:4px!important;
  min-width:0!important;
  flex:1!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-list-item-title{
  display:flex!important;
  align-items:center!important;
  gap:8px!important;
  min-width:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-list-item-title strong{
  font-size:12px!important;
  line-height:1.35!important;
  word-break:break-word!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-list-badge{
  display:inline-flex!important;
  align-items:center!important;
  flex-shrink:0!important;
  padding:2px 7px!important;
  border-radius:999px!important;
  font-size:10px!important;
  font-weight:700!important;
  letter-spacing:.02em!important;
  text-transform:uppercase!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-list-badge.library{
  background:rgba(55,174,28,.16)!important;
  color:#86efac!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-list-badge.list{
  background:rgba(96,165,250,.16)!important;
  color:#93c5fd!important;
}
#${LISTS_LAUNCHER_PANEL_ID}.sp-toolkit-lists-panel-light .sp-toolkit-report-list-badge.library{
  background:#e8f7e4!important;
  color:#2d8d1b!important;
}
#${LISTS_LAUNCHER_PANEL_ID}.sp-toolkit-lists-panel-light .sp-toolkit-report-list-badge.list{
  background:#eef3ff!important;
  color:#3d4fd6!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-list-meta{
  font-size:11px!important;
  line-height:1.35!important;
  opacity:.78!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-list-item.sp-toolkit-report-list-hidden,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-lists-site-heading.sp-toolkit-report-list-hidden,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-tree-node.sp-toolkit-report-list-hidden{
  display:none!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-matrix-mode .sp-toolkit-matrix-sites-list{
  flex:1!important;
  min-height:0!important;
  max-height:none!important;
  overflow-y:auto!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-reports-matrix-mode .sp-toolkit-compass-btnrow{
  flex:0 0 auto!important;
  margin-top:10px!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-reports-host.sp-toolkit-reports-settings-active .sp-toolkit-compass-reports-main,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-reports-host.sp-toolkit-reports-picker-active .sp-toolkit-compass-reports-main{
  display:none!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-reports-host.sp-toolkit-reports-settings-active .sp-toolkit-compass-settings{
  flex:1!important;
  min-height:0!important;
  overflow-y:auto!important;
  margin-top:0!important;
  padding-top:0!important;
  padding-bottom:16px!important;
  border-top:none!important;
  -webkit-mask-image:none!important;
  mask-image:none!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-reports-host.sp-toolkit-reports-picker-active .sp-toolkit-compass-picker{
  flex:1!important;
  min-height:0!important;
  overflow:hidden!important;
  display:flex!important;
  flex-direction:column!important;
  margin-top:0!important;
  padding-top:0!important;
  border-top:none!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane[data-compass-pane='reports'] .sp-toolkit-compass-label:first-of-type{
  margin-top:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-matrix-site-item label{display:flex;align-items:center;gap:8px;margin:0;}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-matrix-site-item input{
  width:16px!important;
  height:16px!important;
  min-width:16px!important;
  min-height:16px!important;
  margin:0!important;
  flex-shrink:0!important;
  accent-color:#37ae1c;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane[data-compass-pane='reports'] .sp-toolkit-compass-report-select,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane[data-compass-pane='reports'] .sp-toolkit-compass-format-select{
  margin-top:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-contents-table .sp-toolkit-filter-row th{
  padding:4px 8px!important;
  vertical-align:middle!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-contents-table .sp-toolkit-filter-row input,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-contents-table .sp-toolkit-filter-row select{
  width:100%!important;
  max-width:100%!important;
  height:36px!important;
  min-height:36px!important;
  max-height:36px!important;
  line-height:1.2!important;
  padding:0 10px!important;
  box-sizing:border-box!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-contents-table .sp-toolkit-filter-row select{
  padding-right:30px!important;
}
#${LISTS_LAUNCHER_PANEL_ID},
#${LISTS_LAUNCHER_PANEL_ID} *{
  -webkit-user-select:text;
  user-select:text;
}
#${LISTS_LAUNCHER_PANEL_ID}{
  --sp-toolkit-motion-seconds:0.22s;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-header-tab,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-toolbar-icon-btn,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-sc-sub,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-ql-list a,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-primary,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-secondary{
  transition-duration:var(--sp-toolkit-motion-seconds)!important;
}
#${LISTS_LAUNCHER_PANEL_ID}.sp-toolkit-theme-switching,
#${LISTS_LAUNCHER_PANEL_ID}.sp-toolkit-theme-switching *:not(.dark-toggle-btn):not(.dark-toggle-knob){
  transition:none!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .dark-toggle-btn .dark-toggle-knob{
  transition:transform .25s cubic-bezier(0.4,0,0.2,1)!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-settings-shell{
  border:1px solid var(--border);
  background:var(--surface);
  border-radius:12px;
  padding:14px 16px 16px;
  display:flex;
  flex-direction:column;
  gap:14px;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-settings-shell > .sp-toolkit-compass-settings-row{
  min-height:32px;
  padding:3px 0;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-field .sp-toolkit-compass-settings-range-wrap{
  display:flex;
  align-items:center;
  gap:12px;
  min-height:36px;
  padding-top:2px;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-settings-row{
  display:flex;
  align-items:center;
  justify-content:space-between;
  gap:14px;
  margin:0;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-settings-row:last-of-type{
  margin-bottom:0;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-settings-row input[type=checkbox]{
  width:16px!important;
  height:16px!important;
  min-width:16px!important;
  min-height:16px!important;
  margin:0!important;
  flex-shrink:0!important;
  padding:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-settings-row label{
  flex:1 1 auto;
  min-width:0;
  font-size:12px;
  line-height:1.45;
  color:var(--txt);
  padding-right:8px;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-settings-shortcuts{
  border:1px dashed var(--border);
  border-radius:10px;
  padding:12px 14px;
  margin:0;
  display:flex;
  flex-direction:column;
  gap:10px;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-settings-shortcuts .sp-toolkit-compass-settings-row{
  min-height:40px;
  padding:4px 0;
  gap:16px;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-settings-shortcuts input[type="text"]{
  width:168px!important;
  max-width:48%!important;
  flex:0 0 auto!important;
  min-height:36px!important;
  height:36px!important;
  padding:8px 12px!important;
  text-align:center;
  text-transform:lowercase;
  font-family:ui-monospace, SFMono-Regular, Menlo, Consolas, monospace!important;
  font-size:11px!important;
  letter-spacing:.02em;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-settings-range-wrap{
  display:flex;
  align-items:center;
  gap:12px;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-settings-range{
  flex:1;
  min-width:0;
  padding:0!important;
  min-height:auto!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-settings-range-value{
  min-width:118px;
  flex-shrink:0;
  text-align:right;
  color:var(--txt-mid);
  font-size:12px;
  font-variant-numeric:tabular-nums;
  line-height:1.35;
  padding-left:4px;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane[data-compass-pane='toolkitSettings'] .sp-toolkit-compass-scroll{
  padding-bottom:18px!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-settings-host{
  padding-bottom:4px;
}
/* ── Canonical form + table layout (single source of truth) ── */
#${LISTS_LAUNCHER_PANEL_ID}{
  --compass-panel-x:14px;
  --compass-field-gap:8px;
  --compass-stack-gap:14px;
  --compass-table-x:12px;
  --compass-list-toolbar-gap:8px;
  --compass-list-toolbar-pad-bottom:10px;
  --compass-list-count-pad-bottom:8px;
  --compass-list-header-pad-y:8px;
  --compass-list-header-pad-x:12px;
  --compass-list-row-pad-y:8px;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-scroll{
  padding:14px var(--compass-panel-x) 16px!important;
}
/* Bottom scroll fade — inner scrollports only (not form panels or whole tab scroll) */
#${LISTS_LAUNCHER_PANEL_ID} #contextList,
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaList,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-pane[data-compass-pane='quicklinks'] .sp-toolkit-compass-scroll,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-contents-wrap,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-subsites-view,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-site-contents-stack .sp-toolkit-subsites-view.visible,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-column-list,
#${LISTS_LAUNCHER_PANEL_ID} #refinableMappingsOut,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-export-progress-log,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-matrix-sites-list,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-report-lists-list,
#${LISTS_LAUNCHER_PANEL_ID} .column-creator-site-list{
  -webkit-mask-image:linear-gradient(to bottom,#000 0,#000 calc(100% - 2.75rem),rgba(0,0,0,.65) calc(100% - 1.25rem),transparent 100%);
  mask-image:linear-gradient(to bottom,#000 0,#000 calc(100% - 2.75rem),rgba(0,0,0,.65) calc(100% - 1.25rem),transparent 100%);
  -webkit-mask-size:100% 100%;
  mask-size:100% 100%;
  -webkit-mask-repeat:no-repeat;
  mask-repeat:no-repeat;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-form-stack,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-reports-config{
  display:flex!important;
  flex-direction:column!important;
  align-items:stretch!important;
  gap:var(--compass-stack-gap)!important;
  width:100%!important;
  min-width:0!important;
  margin:0!important;
  padding:0!important;
  box-sizing:border-box!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-report-opts{
  display:none;
  flex-direction:column!important;
  align-items:stretch!important;
  gap:var(--compass-stack-gap)!important;
  width:100%!important;
  min-width:0!important;
  margin:0!important;
  padding:0!important;
  box-sizing:border-box!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-settings.sp-toolkit-compass-settings-open{
  display:flex!important;
  flex-direction:column!important;
  align-items:stretch!important;
  gap:var(--compass-stack-gap)!important;
  width:100%!important;
  min-width:0!important;
  margin:0!important;
  padding:0!important;
  box-sizing:border-box!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-field,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-matrix-field{
  display:flex!important;
  flex-direction:column!important;
  align-items:stretch!important;
  gap:var(--compass-field-gap)!important;
  width:100%!important;
  min-width:0!important;
  margin:0!important;
  padding:0!important;
  box-sizing:border-box!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-label,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-matrix-field-label,
#${LISTS_LAUNCHER_PANEL_ID} label.sp-toolkit-compass-label,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-field>label,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-matrix-field>label{
  display:block!important;
  width:100%!important;
  margin:0!important;
  padding:0!important;
  text-align:left!important;
  box-sizing:border-box!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-field select,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-field input:not([type=checkbox]):not([type=radio]):not([type=range]),
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-matrix-field select,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-matrix-field input{
  width:100%!important;
  max-width:100%!important;
  margin:0!important;
  box-sizing:border-box!important;
}
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-btnrow{
  width:100%!important;
  margin-top:8px!important;
  padding:0!important;
  box-sizing:border-box!important;
  gap:8px!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .schema-col-header,
#${LISTS_LAUNCHER_PANEL_ID} #contextPanel .compass-context-schema-head{
  padding:var(--compass-list-header-pad-y) var(--compass-list-header-pad-x)!important;
  margin:0!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaList .search-schema-row,
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel #searchSchemaList .search-schema-row,
#${LISTS_LAUNCHER_PANEL_ID} #contextList .search-schema-row,
#${LISTS_LAUNCHER_PANEL_ID} #contextList .compass-page-props-row{
  padding:var(--compass-list-row-pad-y) var(--compass-list-header-pad-x)!important;
  align-items:start!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .col-internal,
#${LISTS_LAUNCHER_PANEL_ID} .search-schema-row .col-internal,
#${LISTS_LAUNCHER_PANEL_ID} .sp-toolkit-compass-colrow .col-internal{
  min-width:0!important;
  white-space:normal!important;
  word-break:break-all!important;
  overflow-wrap:anywhere!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .col-crawled,
#${LISTS_LAUNCHER_PANEL_ID} .search-schema-row .col-crawled{
  min-width:0!important;
  white-space:normal!important;
  word-break:break-all!important;
  overflow-wrap:anywhere!important;
  overflow:visible!important;
  text-overflow:clip!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .sp-toolkit-cc-name,
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .col-name-link{
  min-width:0!important;
  word-break:break-word!important;
  overflow-wrap:anywhere!important;
}
#${LISTS_LAUNCHER_PANEL_ID} #searchSchemaPanel .col-type{
  align-self:start!important;
  justify-self:start!important;
}
`;

let listsLauncherExpanded = false;
let listsLauncherListsListener = null;
let listsLauncherSubsitesListener = null;

/** Global user preference (not per site): saved launcher position in chrome.storage.local. */
const LISTS_LAUNCHER_POSITION_KEY = "listsLauncherPosition";
const LISTS_LAUNCHER_ICON_SCALE_KEY = "listsLauncherIconScale";
const LAUNCHER_BTN_PX_DEFAULT = 58;
/** Matches compassToolkitPanels DEFAULT_LAUNCHER_ICON_SCALE — used when storage has no saved scale. */
const DEFAULT_LAUNCHER_ICON_SCALE = 1.28;
const LAUNCHER_BTN_SCALE_MIN = 0.75;
const LAUNCHER_BTN_SCALE_MAX = 1.6;
let launcherBtnPx = LAUNCHER_BTN_PX_DEFAULT;
const LAUNCHER_VIEW_MARGIN = 14;
const LAUNCHER_DRAG_START_SLOP = 10;
/** 0 = drag as soon as movement exceeds slop while pressed (almost instant). */
const LAUNCHER_DRAG_HOLD_MS = 0;
/** Release within this duration and with little movement counts as a tap (open panel). */
const LAUNCHER_TAP_MAX_MS = 300;
const LISTS_LAUNCHER_ANIMATION_SECONDS_KEY = "listsLauncherAnimationSeconds";
const LISTS_LAUNCHER_ANIMATION_DEFAULT_SECONDS = 0.22;
const LISTS_LAUNCHER_ANIMATION_MIN_SECONDS = 0.06;
const LISTS_LAUNCHER_ANIMATION_MAX_SECONDS = 0.9;
const COMPASS_SHORTCUTS_ENABLED_KEY = "compassShortcutMacrosEnabled";
const COMPASS_SHORTCUTS_KEY = "compassShortcutMacros";
const DEFAULT_COMPASS_SHORTCUTS = {
  openPanel: "alt+shift+s",
  openSearch: "alt+shift+f",
  quicklinks: "alt+shift+1",
  context: "alt+shift+2",
  columns: "alt+shift+3",
  reports: "alt+shift+4",
  refinableProps: "alt+shift+5",
  viewManager: "alt+shift+6",
  siteContents: "alt+shift+7",
  toolkitSettings: "alt+shift+8"
};
let compassShortcutState = {
  enabled: false,
  map: Object.assign({}, DEFAULT_COMPASS_SHORTCUTS)
};

function normalizeLauncherIconScale(value) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) return DEFAULT_LAUNCHER_ICON_SCALE;
  return Math.min(LAUNCHER_BTN_SCALE_MAX, Math.max(LAUNCHER_BTN_SCALE_MIN, parsed));
}

function playLauncherTapeSpin(container, extending) {
  if (!container) return;
  const img = container.querySelector(".sp-toolkit-launcher-compass");
  if (!img) return;
  img.classList.remove("sp-toolkit-launcher-compass-tape-cw", "sp-toolkit-launcher-compass-tape-ccw");
  void img.offsetWidth;
  img.classList.add(extending ? "sp-toolkit-launcher-compass-tape-cw" : "sp-toolkit-launcher-compass-tape-ccw");
  function cleanup() {
    img.removeEventListener("animationend", cleanup);
    img.classList.remove("sp-toolkit-launcher-compass-tape-cw", "sp-toolkit-launcher-compass-tape-ccw");
  }
  img.addEventListener("animationend", cleanup);
}

function applyLauncherIconScale(container, scaleValue) {
  if (!container) return;
  const scale = normalizeLauncherIconScale(scaleValue);
  launcherBtnPx = Math.round(LAUNCHER_BTN_PX_DEFAULT * scale);
  container.style.setProperty("--sp-toolkit-launcher-size", launcherBtnPx + "px");
  container.style.setProperty("--sp-toolkit-launcher-icon-size", Math.max(20, Math.round(launcherBtnPx * 0.52)) + "px");
  const rect = container.getBoundingClientRect();
  const clamped = clampLauncherPosition(rect.left, rect.top);
  container.style.left = clamped.left + "px";
  container.style.top = clamped.top + "px";
}

function normalizeLauncherAnimationSeconds(value) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) return LISTS_LAUNCHER_ANIMATION_DEFAULT_SECONDS;
  return Math.min(LISTS_LAUNCHER_ANIMATION_MAX_SECONDS, Math.max(LISTS_LAUNCHER_ANIMATION_MIN_SECONDS, parsed));
}

function applyListsPanelAnimationSpeed(panel, seconds) {
  if (!panel) return;
  const normalized = normalizeLauncherAnimationSeconds(seconds);
  panel.style.setProperty("--sp-toolkit-motion-seconds", normalized.toFixed(2) + "s");
}

function normalizeShortcutString(raw) {
  return String(raw || "")
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/control/g, "ctrl")
    .replace(/command/g, "meta")
    .replace(/option/g, "alt");
}

function eventShortcutString(e) {
  const key = String(e.key || "").toLowerCase();
  if (!key) return "";
  const parts = [];
  if (e.ctrlKey) parts.push("ctrl");
  if (e.altKey) parts.push("alt");
  if (e.shiftKey) parts.push("shift");
  if (e.metaKey) parts.push("meta");
  if (["control", "alt", "shift", "meta"].indexOf(key) >= 0) return "";
  parts.push(key);
  return parts.join("+");
}

function isTypingTarget(target) {
  if (!target) return false;
  const tag = (target.tagName || "").toLowerCase();
  if (tag === "input" || tag === "textarea" || tag === "select") return true;
  if (target.isContentEditable) return true;
  const role = target.getAttribute && target.getAttribute("role");
  return role === "textbox";
}

function refreshCompassShortcutState() {
  chrome.storage.local.get([COMPASS_SHORTCUTS_ENABLED_KEY, COMPASS_SHORTCUTS_KEY], function (st) {
    compassShortcutState.enabled = !!st[COMPASS_SHORTCUTS_ENABLED_KEY];
    const merged = Object.assign({}, DEFAULT_COMPASS_SHORTCUTS, st[COMPASS_SHORTCUTS_KEY] || {});
    Object.keys(merged).forEach(function (k) {
      merged[k] = normalizeShortcutString(merged[k]);
    });
    compassShortcutState.map = merged;
  });
}

function openCompassPanelAndRun(runFn) {
  const panel = document.getElementById(LISTS_LAUNCHER_PANEL_ID);
  if (!panel) return;
  if (!listsLauncherExpanded) toggleListsLauncher();
  let tries = 0;
  function attempt() {
    tries += 1;
    try {
      if (runFn()) return;
    } catch (_) {}
    if (tries < 14) setTimeout(attempt, 70);
  }
  setTimeout(attempt, 40);
}

function handleCompassShortcutAction(action) {
  if (!action) return;
  if (action === "openPanel") {
    const panel = document.getElementById(LISTS_LAUNCHER_PANEL_ID);
    if (!panel) return;
    if (!listsLauncherExpanded) toggleListsLauncher();
    return;
  }
  if (action === "openSearch") {
    openCompassPanelAndRun(function () {
      if (typeof window.SPOT_compassOpenUniversalSearch !== "function") return false;
      window.SPOT_compassOpenUniversalSearch();
      return true;
    });
    return;
  }
  openCompassPanelAndRun(function () {
    if (typeof window.SPOT_compassNavigateToPane !== "function") return false;
    window.SPOT_compassNavigateToPane(action);
    return true;
  });
}

function bindCompassShortcutMacros() {
  if (window._SPOT_compassShortcutMacrosBound) return;
  window._SPOT_compassShortcutMacrosBound = true;
  document.addEventListener(
    "keydown",
    function (e) {
      if (!compassShortcutState.enabled) return;
      if (isTypingTarget(e.target)) return;
      const typed = eventShortcutString(e);
      if (!typed) return;
      const map = compassShortcutState.map || {};
      const hit = Object.keys(map).find(function (k) {
        return map[k] && map[k] === typed;
      });
      if (!hit) return;
      e.preventDefault();
      e.stopPropagation();
      handleCompassShortcutAction(hit);
    },
    true
  );
}

function clampLauncherPosition(left, top) {
  const maxL = Math.max(LAUNCHER_VIEW_MARGIN, window.innerWidth - launcherBtnPx - LAUNCHER_VIEW_MARGIN);
  const maxT = Math.max(LAUNCHER_VIEW_MARGIN, window.innerHeight - launcherBtnPx - LAUNCHER_VIEW_MARGIN);
  return {
    left: Math.min(maxL, Math.max(LAUNCHER_VIEW_MARGIN, left)),
    top: Math.min(maxT, Math.max(LAUNCHER_VIEW_MARGIN, top)),
  };
}

function updateLauncherPanelPlacement(container) {
  if (!container) return;
  const rect = container.getBoundingClientRect();
  const vw = window.innerWidth;
  const vh = window.innerHeight;
  const cx = rect.left + rect.width / 2;
  const cy = rect.top + rect.height / 2;
  const openDown = cy < vh / 2;
  const anchorLeft = cx < vw / 2;
  container.classList.remove(
    "sp-toolkit-launcher-panel-up",
    "sp-toolkit-launcher-panel-down",
    "sp-toolkit-launcher-panel-anchor-left",
    "sp-toolkit-launcher-panel-anchor-right"
  );
  if (openDown) container.classList.add("sp-toolkit-launcher-panel-down");
  else container.classList.add("sp-toolkit-launcher-panel-up");
  if (anchorLeft) container.classList.add("sp-toolkit-launcher-panel-anchor-left");
  else container.classList.add("sp-toolkit-launcher-panel-anchor-right");
  const panel = document.getElementById(LISTS_LAUNCHER_PANEL_ID);
  if (panel && typeof window.SPOT_updateCompassTabIndicator === "function") {
    window.SPOT_updateCompassTabIndicator(panel, true);
  }
}

function applyLauncherPositionFromStored(container, storedPos) {
  if (!container) return;
  if (
    storedPos &&
    typeof storedPos.left === "number" &&
    typeof storedPos.top === "number" &&
    !Number.isNaN(storedPos.left) &&
    !Number.isNaN(storedPos.top)
  ) {
    const c = clampLauncherPosition(storedPos.left, storedPos.top);
    container.style.left = c.left + "px";
    container.style.top = c.top + "px";
    container.style.right = "auto";
    container.style.bottom = "auto";
  } else {
    container.style.left = "auto";
    container.style.top = "auto";
    container.style.right = LAUNCHER_VIEW_MARGIN + "px";
    container.style.bottom = LAUNCHER_VIEW_MARGIN + "px";
  }
  updateLauncherPanelPlacement(container);
}

function persistLauncherPosition(left, top) {
  const c = clampLauncherPosition(left, top);
  try {
    chrome.storage.local.set({ [LISTS_LAUNCHER_POSITION_KEY]: { left: c.left, top: c.top } });
  } catch (_) {}
}

let listsLauncherResizeTimer = null;
function bindListsLauncherResizeClamp(container) {
  function clampStoredToViewport() {
    if (!container || !document.body.contains(container)) return;
    chrome.storage.local.get(LISTS_LAUNCHER_POSITION_KEY, function (r) {
      const p = r && r[LISTS_LAUNCHER_POSITION_KEY];
      applyLauncherPositionFromStored(container, p);
    });
  }
  window.addEventListener("resize", function () {
    if (listsLauncherResizeTimer) clearTimeout(listsLauncherResizeTimer);
    listsLauncherResizeTimer = setTimeout(clampStoredToViewport, 120);
  });
}

/** Press and hold, then move, to drag. Hovering without button down does nothing. Release saves position. */
function attachListsLauncherDragAndTap(container, btn) {
  let pointerDownTime = 0;
  let startClientX = 0;
  let startClientY = 0;
  let dragging = false;
  let dragOffsetX = 0;
  let dragOffsetY = 0;
  let suppressNextClick = false;
  /** Non-null only while primary pointer is down on the button (prevents hover-only moves from dragging). */
  let activePointerId = null;

  function moveDist(ev) {
    return Math.hypot(ev.clientX - startClientX, ev.clientY - startClientY);
  }

  function onPointerDown(ev) {
    if (ev.button != null && ev.button !== 0) return;
    pointerDownTime = Date.now();
    startClientX = ev.clientX;
    startClientY = ev.clientY;
    dragging = false;
    activePointerId = ev.pointerId;
    try {
      if (btn.setPointerCapture) btn.setPointerCapture(ev.pointerId);
    } catch (_) {}
  }

  function onPointerMove(ev) {
    if (activePointerId == null || ev.pointerId !== activePointerId) return;
    if (dragging) {
      const nextLeft = ev.clientX - dragOffsetX;
      const nextTop = ev.clientY - dragOffsetY;
      const c = clampLauncherPosition(nextLeft, nextTop);
      container.style.left = c.left + "px";
      container.style.top = c.top + "px";
      container.style.right = "auto";
      container.style.bottom = "auto";
      updateLauncherPanelPlacement(container);
      return;
    }
    const held = Date.now() - pointerDownTime;
    if (held < LAUNCHER_DRAG_HOLD_MS) return;
    if (moveDist(ev) <= LAUNCHER_DRAG_START_SLOP) return;
    dragging = true;
    container.classList.add("sp-toolkit-launcher-dragging");
    const r = container.getBoundingClientRect();
    dragOffsetX = ev.clientX - r.left;
    dragOffsetY = ev.clientY - r.top;
    const c = clampLauncherPosition(ev.clientX - dragOffsetX, ev.clientY - dragOffsetY);
    container.style.left = c.left + "px";
    container.style.top = c.top + "px";
    container.style.right = "auto";
    container.style.bottom = "auto";
    updateLauncherPanelPlacement(container);
  }

  function onPointerUp(ev) {
    if (activePointerId == null || ev.pointerId !== activePointerId) return;
    try {
      if (btn.releasePointerCapture) btn.releasePointerCapture(activePointerId);
    } catch (_) {}
    activePointerId = null;
    const distEnd = Math.hypot(ev.clientX - startClientX, ev.clientY - startClientY);
    const elapsed = Date.now() - pointerDownTime;
    if (dragging) {
      dragging = false;
      container.classList.remove("sp-toolkit-launcher-dragging");
      const r = container.getBoundingClientRect();
      persistLauncherPosition(r.left, r.top);
      suppressNextClick = true;
      updateLauncherPanelPlacement(container);
      return;
    }
    if (elapsed <= LAUNCHER_TAP_MAX_MS && distEnd <= LAUNCHER_DRAG_START_SLOP) {
      suppressNextClick = true;
      toggleListsLauncher();
    }
  }

  function onPointerCancel(ev) {
    if (activePointerId == null || (ev && ev.pointerId != null && ev.pointerId !== activePointerId)) return;
    try {
      if (btn.releasePointerCapture) btn.releasePointerCapture(activePointerId);
    } catch (_) {}
    activePointerId = null;
    if (dragging) {
      dragging = false;
      container.classList.remove("sp-toolkit-launcher-dragging");
      const r = container.getBoundingClientRect();
      persistLauncherPosition(r.left, r.top);
      suppressNextClick = true;
      updateLauncherPanelPlacement(container);
    }
  }

  btn.addEventListener("pointerdown", onPointerDown);
  btn.addEventListener("pointermove", onPointerMove);
  btn.addEventListener("pointerup", onPointerUp);
  btn.addEventListener("pointercancel", onPointerCancel);
  btn.addEventListener(
    "click",
    function (e) {
      if (!suppressNextClick) return;
      e.preventDefault();
      e.stopImmediatePropagation();
      suppressNextClick = false;
    },
    true
  );
}

const siteListsCacheBySite = Object.create(null);
const siteSubsitesCacheBySite = Object.create(null);
let pageContextFullCacheHref = "";
let pageContextFullCacheData = null;

function getCompassSiteKeyFromLocation() {
  try {
    const path = decodeURIComponent((location.pathname || "").replace(/\/$/, "") || "");
    const segments = path.split("/").filter(Boolean);
    let siteEnd = -1;
    for (let i = 0; i < segments.length; i++) {
      if (["sites", "site", "teams"].indexOf(segments[i]) >= 0 && segments[i + 1]) {
        siteEnd = i + 1;
        break;
      }
    }
    return siteEnd >= 0 ? location.origin + "/" + segments.slice(0, siteEnd + 1).join("/") : location.origin;
  } catch (_) {
    return location.origin || "";
  }
}

function resetCompassPanelForSiteChange(panel) {
  const loadingEl = panel.querySelector(".sp-toolkit-lists-loading");
  const errorEl = panel.querySelector(".sp-toolkit-lists-error");
  const tableEl = panel.querySelector(".sp-toolkit-contents-table");
  const contentsWrap = panel.querySelector(".sp-toolkit-contents-wrap");
  if (loadingEl) loadingEl.style.display = "flex";
  if (errorEl) errorEl.style.display = "none";
  if (contentsWrap) contentsWrap.style.display = "block";
  if (tableEl) {
    tableEl.classList.remove("sp-toolkit-lists-loaded");
    tableEl.style.display = "none";
    const tbody = tableEl.querySelector("tbody");
    if (tbody) tbody.innerHTML = "";
  }
  const prevSubsites = contentsWrap ? contentsWrap.querySelector(".sp-toolkit-subsites-block") : null;
  if (prevSubsites) prevSubsites.remove();
  try {
    delete panel._spToolkitSubsitesCache;
    delete panel._siteContentsSubsites;
    delete panel._siteContentsRenderRows;
  } catch (_) {}
  const navTabs = panel.querySelector(".sp-toolkit-compass-nav-tabs");
  if (navTabs) {
    navTabs.innerHTML = "";
    delete navTabs.dataset.compassTabsBuilt;
  }
  try {
    delete panel.dataset.compassToolListeners;
    delete panel.dataset.compassInitialized;
    delete panel.dataset.compassQuicklinksReady;
    delete panel.dataset.compassCurrentListId;
  } catch (_) {}
  const qh = panel.querySelector(".sp-toolkit-compass-quicklinks-host");
  if (qh) qh.innerHTML = "";
  const colsList = panel.querySelector("#searchSchemaList");
  if (colsList) colsList.innerHTML = "";
  const colsErr = panel.querySelector(".sp-toolkit-compass-columns-error");
  if (colsErr) {
    colsErr.textContent = "";
    colsErr.style.display = "none";
  }
  const repHost = panel.querySelector(".sp-toolkit-compass-reports-host");
  if (repHost) {
    repHost.innerHTML = "";
    delete repHost.dataset.built;
    delete repHost.dataset.building;
  }
  const refinOut = panel.querySelector(".sp-toolkit-compass-refinable-out");
  if (refinOut) {
    refinOut.textContent = "Select a refinable, then mappings load when you open this tab.";
    refinOut.className = "sp-toolkit-compass-refinable-out sp-toolkit-compass-muted";
  }
  const typeSel = panel.querySelector(".sp-toolkit-compass-refinable-type");
  const numSel = panel.querySelector(".sp-toolkit-compass-refinable-num");
  if (typeSel) typeSel.innerHTML = "";
  if (numSel) numSel.innerHTML = "";
    delete panel.dataset.refinableControlsFilled;
    delete panel.dataset.compassPrefetchStarted;
    delete panel.dataset.compassQuicklinksLoading;
    delete panel.dataset.compassContextLoading;
    delete panel.dataset.compassColumnsLoading;
    delete panel.dataset.compassRefinableLoading;
    delete panel._compassPaneRegistry;
    delete panel.dataset.compassActivePane;
    delete panel.dataset.compassTitlePane;
  const vh = panel.querySelector(".sp-toolkit-compass-views-host");
  if (vh) {
    vh.innerHTML = "";
    delete vh.dataset.wired;
  }
  const titleEl = panel.querySelector(".sp-toolkit-panel-title-label");
  if (titleEl) titleEl.innerHTML = "";
  try {
    delete panel._compassTitlePrefixHtml;
  } catch (_) {}
  panel.querySelectorAll(".sp-toolkit-compass-pane").forEach(function (p) {
    p.classList.toggle("sp-toolkit-compass-pane-active", p.getAttribute("data-compass-pane") === "siteContents");
  });
}

function initCompassPanelEarly(panel, siteKey, currentListId) {
  const onSiteContentsPage = /viewlsts\.aspx/i.test(location.pathname || "");
  function finish(isListPage) {
    if (typeof window.SPOT_initCompassToolkitPanels !== "function") return;
    window.SPOT_initCompassToolkitPanels(panel, toolkitActionPromise, {
      siteUrl: siteKey,
      onSiteContentsPage: onSiteContentsPage,
      onListOrLibraryView: !!isListPage,
      currentListId: currentListId || ""
    });
    panel.dataset.compassInitialized = "1";
    panel.dataset.compassSiteKey = siteKey;
    if (currentListId) panel.dataset.compassCurrentListId = currentListId;
    if (typeof window.SPOT_noteCompassPanelContentChanged === "function") {
      window.SPOT_noteCompassPanelContentChanged(panel);
    }
  }
  toolkitActionPromise({ action: "checkListPage" }).then(function (chk) {
    finish(!!(chk && chk.isListPage));
  }).catch(function () {
    finish(false);
  });
}

function applySiteContentsSubsites(panel, subsites) {
  if (!panel) return;
  panel._spToolkitSubsitesCache = Array.isArray(subsites) ? subsites : [];
  panel._siteContentsSubsites = panel._spToolkitSubsitesCache.map(function (s) {
    return {
      isSubsite: true,
      title: s.title || "Untitled",
      viewUrl: s.url || "",
      typeLabel: "Subsite",
      itemCount: null,
      modified: null
    };
  });
  const typeSelect = panel.querySelector(".sp-toolkit-contents-table .sp-toolkit-filter-row select");
  if (typeSelect && !typeSelect.querySelector('option[value="Subsite"]')) {
    const opt = document.createElement("option");
    opt.value = "Subsite";
    opt.textContent = "Subsite";
    typeSelect.appendChild(opt);
  }
  if (typeof panel._siteContentsRenderRows === "function") {
    panel._siteContentsRenderRows();
  }
}

function toggleListsLauncher() {
  const container = document.getElementById(LISTS_LAUNCHER_ID);
  const panel = document.getElementById(LISTS_LAUNCHER_PANEL_ID);
  if (!container || !panel) return;
  listsLauncherExpanded = !listsLauncherExpanded;
  if (listsLauncherExpanded) {
    if (typeof window.SPOT_resetCompassPanelSize === "function") {
      window.SPOT_resetCompassPanelSize(panel);
    } else {
      panel.style.width = "";
      panel.style.height = "";
    }
    panel.style.display = "flex";
    panel.classList.remove("sp-toolkit-panel-rolling-up");
    playLauncherTapeSpin(container, true);
    panel.classList.add("sp-toolkit-lists-panel-open");
    updateLauncherPanelPlacement(container);
    try { chrome.runtime.sendMessage({ type: "SPOToolkitPanelOpened" }); } catch (_) {}
    startCompassUrlWatch();
  } else {
    stopCompassUrlWatch();
    playLauncherTapeSpin(container, false);
    panel.classList.remove("sp-toolkit-lists-panel-open");
    panel.classList.add("sp-toolkit-panel-rolling-up");
    setTimeout(function () {
      panel.style.display = "none";
      panel.classList.remove("sp-toolkit-panel-rolling-up");
    }, 130);
  }
  if (listsLauncherExpanded) {
    const compassSiteKey = getCompassSiteKeyFromLocation();
    const siteChanged = !!(panel.dataset.compassSiteKey && panel.dataset.compassSiteKey !== compassSiteKey);
    const canWarmReopen =
      !siteChanged &&
      panel.dataset.compassInitialized === "1" &&
      !!panel.querySelector(".sp-toolkit-compass-nav-tabs .sp-toolkit-compass-pane-tab");
    if (canWarmReopen) {
      const onSiteContentsPage = /viewlsts\.aspx/i.test(location.pathname || "");
      toolkitActionPromise({ action: "checkListPage" }).then(function (chk) {
        if (typeof window.SPOT_applyCompassTabAvailability === "function") {
          window.SPOT_applyCompassTabAvailability(panel, !!(chk && chk.isListPage));
        }
        if (typeof window.SPOT_reopenCompassPanel === "function") {
          window.SPOT_reopenCompassPanel(panel, toolkitActionPromise, {
            siteUrl: compassSiteKey,
            onSiteContentsPage: onSiteContentsPage,
            onListOrLibraryView: !!(chk && chk.isListPage),
            currentListId: panel.dataset.compassCurrentListId || ""
          });
        }
      }).catch(function () {
        if (typeof window.SPOT_reopenCompassPanel === "function") {
          window.SPOT_reopenCompassPanel(panel, toolkitActionPromise, {
            siteUrl: compassSiteKey,
            onSiteContentsPage: onSiteContentsPage,
            onListOrLibraryView: false,
            currentListId: panel.dataset.compassCurrentListId || ""
          });
        }
      });
      return;
    }
    if (siteChanged || panel.dataset.compassInitialized !== "1") {
      resetCompassPanelForSiteChange(panel);
    }
    panel.dataset.compassSiteKey = compassSiteKey;
    if (listsLauncherListsListener) {
      window.removeEventListener("message", listsLauncherListsListener);
      listsLauncherListsListener = null;
    }
    if (listsLauncherSubsitesListener) {
      window.removeEventListener("message", listsLauncherSubsitesListener);
      listsLauncherSubsitesListener = null;
    }
    initCompassPanelEarly(panel, compassSiteKey, "");
    listsLauncherListsListener = (ev) => {
      if (!ev.data || ev.data.__spcsv !== true || ev.data.type !== "SPCSVSiteListsResult") return;
      window.removeEventListener("message", listsLauncherListsListener);
      listsLauncherListsListener = null;
      const tableEl = panel.querySelector(".sp-toolkit-contents-table");
      const tbody = tableEl ? tableEl.querySelector("tbody") : null;
      const loadingEl = panel.querySelector(".sp-toolkit-lists-loading");
      const errorEl = panel.querySelector(".sp-toolkit-lists-error");
      const contentsWrap = panel.querySelector(".sp-toolkit-contents-wrap");
      if (loadingEl) loadingEl.style.display = "none";
      if (contentsWrap) contentsWrap.style.display = "block";
      if (ev.data.error) {
        if (errorEl) {
          errorEl.textContent = ev.data.error;
          errorEl.style.display = "block";
        }
        return;
      }
      if (errorEl) errorEl.style.display = "none";
      if (tableEl && !ev.data.error) {
        tableEl.classList.add("sp-toolkit-lists-loaded");
        tableEl.style.display = "table";
      }
      const siteUrl = (ev.data.siteUrl || "").replace(/\/$/, "");
      /* Fluent Key 24 regular — Fabric MDL2 "Permissions" / Unicode E8D7 (flicon.io) */
      const fluentPermissionsKeySvgInner = "<path fill=\"currentColor\" d=\"M18.2493 6.99982C18.2493 7.69017 17.6896 8.24982 16.9993 8.24982C16.3089 8.24982 15.7493 7.69017 15.7493 6.99982C15.7493 6.30946 16.3089 5.74982 16.9993 5.74982C17.6896 5.74982 18.2493 6.30946 18.2493 6.99982ZM15.4992 2.0498C11.885 2.0498 8.94922 4.98559 8.94922 8.5998C8.94922 8.98701 8.99939 9.36017 9.05968 9.70382C9.07749 9.80529 9.04493 9.89344 8.99046 9.94791L2.75467 16.1837C2.23895 16.6994 1.94922 17.3989 1.94922 18.1282V20.2998C1.94922 21.2663 2.73272 22.0498 3.69922 22.0498H6.19922C7.16572 22.0498 7.94922 21.2663 7.94922 20.2998V19.0498H9.69922C10.3896 19.0498 10.9492 18.4902 10.9492 17.7998V16.0498H12.6992C13.3741 16.0498 13.9241 15.515 13.9484 14.846C14.4451 14.9738 14.9689 15.0498 15.4992 15.0498C19.1134 15.0498 22.0492 12.114 22.0492 8.4998C22.0492 4.86866 19.0963 2.0498 15.4992 2.0498ZM10.4492 8.5998C10.4492 5.81402 12.7134 3.5498 15.4992 3.5498C18.3021 3.5498 20.5492 5.73095 20.5492 8.4998C20.5492 11.2856 18.285 13.5498 15.4992 13.5498C14.8199 13.5498 14.1206 13.3787 13.4947 13.1104C13.2629 13.0111 12.9968 13.0349 12.7864 13.1737C12.5759 13.3125 12.4492 13.5477 12.4492 13.7998V14.5498H10.6992C10.0089 14.5498 9.44922 15.1094 9.44922 15.7998V17.5498H7.69922C7.00886 17.5498 6.44922 18.1094 6.44922 18.7998V20.2998C6.44922 20.4379 6.33729 20.5498 6.19922 20.5498H3.69922C3.56115 20.5498 3.44922 20.4379 3.44922 20.2998V18.1282C3.44922 17.7967 3.58091 17.4788 3.81534 17.2443L10.0511 11.0086C10.4695 10.5902 10.6349 10.0018 10.5371 9.44461C10.4834 9.13865 10.4492 8.8622 10.4492 8.5998Z\"/>";
      const lists = Array.isArray(ev.data.lists) ? ev.data.lists : [];
      const currentPath = (location.pathname || "").split("?")[0].replace(/\/$/, "") || "/";
      const onSiteContentsPage = /viewlsts\.aspx/i.test(location.pathname || "");
      let currentListTitle = null;
      for (let i = 0; i < lists.length; i++) {
        const viewPath = lists[i].viewUrl ? (function () { try { return new URL(lists[i].viewUrl).pathname.replace(/\/$/, "") || "/"; } catch (_) { return ""; } }()) : "";
        if (viewPath && (currentPath === viewPath || currentPath.startsWith(viewPath + "/"))) {
          currentListTitle = lists[i].title || "Untitled";
          break;
        }
      }
      function formatModified(s) {
        if (!s) return "";
        try {
          const d = new Date(s);
          if (isNaN(d.getTime())) return String(s);
          const m = d.getMonth() + 1, day = d.getDate(), y = d.getFullYear();
          const h = d.getHours(), min = d.getMinutes(), am = h < 12;
          const h12 = (h % 12) || 12;
          return m + "/" + day + "/" + y + " " + h12 + ":" + (min < 10 ? "0" : "") + min + " " + (am ? "AM" : "PM");
        } catch (_) { return String(s); }
      }
      const iconLibrary = "<svg class=\"sp-toolkit-row-icon\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"><path d=\"M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z\"/><line x1=\"12\" y1=\"11\" x2=\"12\" y2=\"17\"/><line x1=\"9\" y1=\"14\" x2=\"15\" y2=\"14\"/></svg>";
      const iconList = "<svg class=\"sp-toolkit-row-icon\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"><rect x=\"3\" y=\"4\" width=\"18\" height=\"18\" rx=\"2\" ry=\"2\"/><line x1=\"16\" y1=\"2\" x2=\"16\" y2=\"6\"/><line x1=\"8\" y1=\"2\" x2=\"8\" y2=\"6\"/><line x1=\"3\" y1=\"10\" x2=\"21\" y2=\"10\"/></svg>";
      const iconSubsite = "<svg class=\"sp-toolkit-row-icon\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"><path d=\"M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z\"/><polyline points=\"9 22 9 12 15 12 15 22\"/></svg>";
      const thead = tableEl ? tableEl.querySelector("thead") : null;
      if (thead && tableEl) {
        thead.innerHTML = "";
        const cols = [{ key: "name", label: "Name" }, { key: "type", label: "Type" }, { key: "items", label: "Items" }, { key: "modified", label: "Modified" }, { key: "settings", label: "" }];
        const filterRow = document.createElement("tr");
        filterRow.className = "sp-toolkit-filter-row";
        const filterNameTh = document.createElement("th");
        filterNameTh.setAttribute("data-column", "name");
        const nameFilterWrap = document.createElement("div");
        nameFilterWrap.className = "sp-toolkit-site-contents-name-filter";
        nameFilterWrap.innerHTML =
          "<svg class=\"sp-toolkit-sc-filter-icon\" viewBox=\"0 0 24 24\" aria-hidden=\"true\" focusable=\"false\"><path fill=\"currentColor\" fill-rule=\"evenodd\" clip-rule=\"evenodd\" d=\"M10.5 3.75a6.75 6.75 0 100 13.5 6.75 6.75 0 000-13.5zM2.25 10.5a8.25 8.25 0 1114.59 5.31l3.94 3.94a.75.75 0 11-1.06 1.06l-3.94-3.94A8.25 8.25 0 012.25 10.5z\"/></svg>";
        const nameInput = document.createElement("input");
        nameInput.type = "text";
        nameInput.placeholder = "Filter…";
        nameInput.setAttribute("aria-label", "Filter by name");
        nameFilterWrap.appendChild(nameInput);
        filterNameTh.appendChild(nameFilterWrap);
        filterRow.appendChild(filterNameTh);
        const uniqueTypes = [];
        const typeSet = {};
        lists.forEach(function (item) {
          const tl = item.typeLabel != null ? item.typeLabel : (item.isLibrary ? "Document library" : "List");
          if (!typeSet[tl]) { typeSet[tl] = true; uniqueTypes.push(tl); }
        });
        const cachedSubs = panel._siteContentsSubsites || [];
        if (cachedSubs.length && !typeSet.Subsite) {
          typeSet.Subsite = true;
          uniqueTypes.push("Subsite");
        }
        uniqueTypes.sort();
        const filterTypeTh = document.createElement("th");
        filterTypeTh.setAttribute("data-column", "type");
        const typeSelect = document.createElement("select");
        typeSelect.setAttribute("aria-label", "Filter by type");
        const typeAll = document.createElement("option");
        typeAll.value = "";
        typeAll.textContent = "All types";
        typeSelect.appendChild(typeAll);
        uniqueTypes.forEach(function (t) {
          const opt = document.createElement("option");
          opt.value = t;
          opt.textContent = t;
          typeSelect.appendChild(opt);
        });
        filterTypeTh.appendChild(typeSelect);
        filterRow.appendChild(filterTypeTh);
        filterRow.appendChild(document.createElement("th"));
        filterRow.appendChild(document.createElement("th"));
        filterRow.appendChild(document.createElement("th"));
        thead.appendChild(filterRow);
        const headerRow = document.createElement("tr");
        headerRow.className = "sp-toolkit-column-header-row";
        cols.forEach(function (c) {
          const th = document.createElement("th");
          th.setAttribute("data-column", c.key);
          th.innerHTML = c.key === "settings" ? c.label : c.label + " <span class=\"sp-toolkit-sort-icon\" aria-hidden=\"true\">↕</span>";
          headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);
        panel._siteContentsLists = lists;
        panel._siteContentsCurrentPath = currentPath;
        panel._siteContentsSort = { col: "type", dir: 1 };
        panel._siteContentsFilter = { name: "", type: "" };
        function renderRows() {
          if (!tbody) return;
          const listData = (panel._siteContentsLists || []).concat(panel._siteContentsSubsites || []);
          const sort = panel._siteContentsSort || { col: null, dir: 1 };
          const filter = panel._siteContentsFilter || { name: "", type: "" };
          let filtered = listData.filter(function (item) {
            const title = (item.title || "Untitled").toLowerCase();
            const typeLabel = item.typeLabel != null ? item.typeLabel : (item.isSubsite ? "Subsite" : (item.isLibrary ? "Document library" : "List"));
            if (filter.name && title.indexOf(filter.name.toLowerCase()) === -1) return false;
            if (filter.type && typeLabel !== filter.type) return false;
            return true;
          });
          const col = sort.col;
          const dir = sort.dir;
          if (col === "name") filtered.sort(function (a, b) { const x = (a.title || "").toLowerCase(), y = (b.title || "").toLowerCase(); return dir * (x < y ? -1 : x > y ? 1 : 0); });
          else if (col === "type") filtered.sort(function (a, b) { const x = (a.typeLabel != null ? a.typeLabel : (a.isSubsite ? "Subsite" : (a.isLibrary ? "Document library" : "List"))); const y = (b.typeLabel != null ? b.typeLabel : (b.isSubsite ? "Subsite" : (b.isLibrary ? "Document library" : "List"))); return dir * (x < y ? -1 : x > y ? 1 : 0); });
          else if (col === "items") filtered.sort(function (a, b) { const x = (a.itemCount != null ? a.itemCount : -1) | 0; const y = (b.itemCount != null ? b.itemCount : -1) | 0; return dir * (x - y); });
          else if (col === "modified") filtered.sort(function (a, b) { const x = new Date(a.modified || 0).getTime(); const y = new Date(b.modified || 0).getTime(); return dir * (x - y); });
          tbody.innerHTML = "";
          const path = panel._siteContentsCurrentPath || "";
          filtered.forEach(function (item) {
            const viewPath = item.viewUrl ? (function () { try { return new URL(item.viewUrl).pathname.replace(/\/$/, "") || "/"; } catch (_) { return ""; } }()) : "";
            let isCurrent = false;
            if (item.isSubsite) {
              isCurrent = viewPath && (path === viewPath || path.startsWith(viewPath + "/"));
            } else {
              isCurrent = viewPath && (path === viewPath || path.startsWith(viewPath + "/"));
            }
            const tr = document.createElement("tr");
            if (isCurrent) tr.classList.add("sp-toolkit-current");
            const nameCell = document.createElement("td");
            const nameDiv = document.createElement("div");
            nameDiv.className = "sp-toolkit-row-name";
            nameDiv.innerHTML = (item.isSubsite ? iconSubsite : (item.isLibrary ? iconLibrary : iconList)) + " ";
            const a = document.createElement("a");
            a.href = item.viewUrl || "#";
            a.textContent = item.title || "Untitled";
            a.addEventListener("click", function (e) { e.preventDefault(); e.stopPropagation(); closeListsLauncherPanel(); if (item.viewUrl) requestAnimationFrame(function () { window.location.href = item.viewUrl; }); });
            nameDiv.appendChild(a);
            nameCell.appendChild(nameDiv);
            tr.appendChild(nameCell);
            const typeLabel = item.typeLabel != null ? item.typeLabel : (item.isSubsite ? "Subsite" : (item.isLibrary ? "Document library" : "List"));
            const typeCell = document.createElement("td");
            typeCell.className = "sp-toolkit-type-cell";
            const pill = document.createElement("span");
            pill.className = "sp-toolkit-type-pill";
            pill.textContent = typeLabel;
            typeCell.appendChild(pill);
            tr.appendChild(typeCell);
            const itemsCell = document.createElement("td");
            itemsCell.textContent = item.itemCount != null ? String(item.itemCount) : "—";
            tr.appendChild(itemsCell);
            const modCell = document.createElement("td");
            modCell.textContent = item.modified ? formatModified(item.modified) : "—";
            tr.appendChild(modCell);
            const settingsCell = document.createElement("td");
            settingsCell.className = "sp-toolkit-row-settings";
            if (!item.isSubsite) {
            const listGuid = item.id != null ? ("{" + String(item.id).replace(/[{}]/g, "").trim().toUpperCase() + "}") : "";
            if (listGuid && siteUrl) {
              const guidNoBraces = String(item.id).replace(/[{}]/g, "").trim().toUpperCase();
              const vmBtn = document.createElement("button");
              vmBtn.type = "button";
              vmBtn.className = "sp-toolkit-row-view-manager";
              vmBtn.title = "View Manager for this list";
              vmBtn.setAttribute("aria-label", "View Manager for " + (item.title || "list"));
              vmBtn.innerHTML = "<svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" stroke-linecap=\"round\" stroke-linejoin=\"round\"><path d=\"M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z\"/><circle cx=\"12\" cy=\"12\" r=\"3\"/></svg>";
              vmBtn.addEventListener("click", function (e) {
                e.preventDefault();
                e.stopPropagation();
                try {
                  chrome.runtime.sendMessage({ type: "SPOToolkitOpenViewManager", listId: guidNoBraces, webUrl: siteUrl }, function (res) {
                    if (chrome.runtime.lastError || (res && res.ok === false)) {
                      console.warn("SPOToolkit: View Manager open failed", chrome.runtime.lastError || res);
                    }
                  });
                } catch (err) {
                  console.warn("SPOToolkit: View Manager open failed", err);
                }
              });
              settingsCell.appendChild(vmBtn);
              const settingsLink = document.createElement("a");
              settingsLink.href = siteUrl + "/_layouts/15/listedit.aspx?List=" + encodeURIComponent(listGuid);
              settingsLink.target = "_blank";
              settingsLink.rel = "noopener";
              settingsLink.title = "List settings";
              settingsLink.setAttribute("aria-label", "Settings for " + (item.title || "list"));
              settingsLink.innerHTML = "<svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" stroke-linecap=\"round\" stroke-linejoin=\"round\"><line x1=\"4\" y1=\"21\" x2=\"4\" y2=\"14\"/><line x1=\"4\" y1=\"10\" x2=\"4\" y2=\"3\"/><line x1=\"12\" y1=\"21\" x2=\"12\" y2=\"12\"/><line x1=\"12\" y1=\"8\" x2=\"12\" y2=\"3\"/><line x1=\"20\" y1=\"21\" x2=\"20\" y2=\"16\"/><line x1=\"20\" y1=\"12\" x2=\"20\" y2=\"3\"/><line x1=\"1\" y1=\"14\" x2=\"7\" y2=\"14\"/><line x1=\"9\" y1=\"8\" x2=\"15\" y2=\"8\"/><line x1=\"17\" y1=\"16\" x2=\"23\" y2=\"16\"/></svg>";
              settingsCell.appendChild(settingsLink);
              const permsLink = document.createElement("a");
              permsLink.href = siteUrl + "/_layouts/15/user.aspx?obj=" + encodeURIComponent(listGuid + ",list");
              permsLink.target = "_blank";
              permsLink.rel = "noopener";
              permsLink.title = "List/library permissions";
              permsLink.setAttribute("aria-label", "Permissions for " + (item.title || "list"));
              permsLink.innerHTML = "<svg viewBox=\"0 0 24 24\" fill=\"none\" xmlns=\"http://www.w3.org/2000/svg\">" + fluentPermissionsKeySvgInner + "</svg>";
              settingsCell.appendChild(permsLink);
            }
            }
            tr.appendChild(settingsCell);
            tbody.appendChild(tr);
          });
          if (typeof window.SPOT_noteCompassPanelContentChanged === "function") {
            window.SPOT_noteCompassPanelContentChanged(panel);
          }
        }
        function updateSortIcons() {
          const sort = panel._siteContentsSort || { col: null, dir: 1 };
          headerRow.querySelectorAll("th[data-column]").forEach(function (th) {
            const key = th.getAttribute("data-column");
            const icon = th.querySelector(".sp-toolkit-sort-icon");
            if (!icon || key === "settings") return;
            if (sort.col !== key) { icon.textContent = "↕"; return; }
            icon.textContent = sort.dir === 1 ? "↑" : "↓";
          });
        }
        headerRow.querySelectorAll("th[data-column]").forEach(function (th) {
          th.addEventListener("click", function () {
            const key = th.getAttribute("data-column");
            if (key === "settings") return;
            const sort = panel._siteContentsSort || { col: null, dir: 1 };
            if (sort.col === key) panel._siteContentsSort = { col: key, dir: -sort.dir };
            else panel._siteContentsSort = { col: key, dir: 1 };
            updateSortIcons();
            renderRows();
          });
        });
        nameInput.addEventListener("input", function () { panel._siteContentsFilter = panel._siteContentsFilter || {}; panel._siteContentsFilter.name = nameInput.value.trim(); renderRows(); });
        typeSelect.addEventListener("change", function () { panel._siteContentsFilter = panel._siteContentsFilter || {}; panel._siteContentsFilter.type = typeSelect.value; renderRows(); });
        panel._siteContentsRenderRows = renderRows;
        renderRows();
        requestAnimationFrame(function () {
          requestAnimationFrame(function () {
            try {
              const activePane = panel.querySelector(".sp-toolkit-compass-pane.sp-toolkit-compass-pane-active");
              if (activePane && activePane.getAttribute("data-compass-pane") === "siteContents" && nameInput) {
                nameInput.focus({ preventScroll: true });
              }
            } catch (_) {}
          });
        });
        updateSortIcons();
      }
      function escapeHtml(str) {
        const div = document.createElement("div");
        div.textContent = str;
        return div.innerHTML;
      }
      function siteContentsTitlePrefixHtml(tenantName, siteName, listTitle) {
        const t = (tenantName || "").trim();
        const s = (siteName || "").trim();
        const list = (listTitle || "").trim();
        if (s && list) return "<span class=\"sp-toolkit-label-site-name\">" + escapeHtml(s) + "</span>" + " | " + escapeHtml(list);
        if (s) return "<span class=\"sp-toolkit-label-site-name\">" + escapeHtml(s) + "</span>";
        if (t) return escapeHtml(t);
        return "";
      }
      const titlePrefixHtml = siteContentsTitlePrefixHtml(ev.data.tenantName, ev.data.siteName, currentListTitle);
      const tenantName = (ev.data.tenantName || "").trim();
      const siteUrlForLabel = (siteUrl || ev.data.siteUrl || "").replace(/\/$/, "");
      const tenantLineText = tenantName && siteUrlForLabel ? tenantName + " - " + siteUrlForLabel : tenantName || siteUrlForLabel || "";
      const headerLeftEl = panel.querySelector(".sp-toolkit-header-left");
      function applyTenantLine(container, text) {
        if (!container) return;
        const lines = container.querySelectorAll(".sp-toolkit-header-tenant-line");
        const first = lines[0];
        if (first) {
          first.textContent = text;
          first.style.display = text ? "" : "none";
          for (let i = 1; i < lines.length; i++) lines[i].remove();
        } else if (text) {
          const el = document.createElement("div");
          el.className = "sp-toolkit-header-tenant-line";
          el.textContent = text;
          el.setAttribute("aria-hidden", "true");
          const navEl = container.querySelector(".sp-toolkit-compass-nav-tabs");
          if (navEl) container.insertBefore(el, navEl);
          else container.appendChild(el);
        }
      }
      panel._compassTitlePrefixHtml = titlePrefixHtml;
      applyTenantLine(headerLeftEl, tenantLineText);
      const currentListIdForPanel = ev.data.currentListId && String(ev.data.currentListId).replace(/^\{|\}$/g, "").trim();
      if (currentListIdForPanel) panel.dataset.compassCurrentListId = currentListIdForPanel;
      const cacheKey = (siteUrl || "").replace(/\/$/, "");
      if (cacheKey) {
        siteListsCacheBySite[cacheKey] = {
          __spcsv: true,
          type: "SPCSVSiteListsResult",
          lists: ev.data.lists,
          siteUrl: siteUrl,
          currentListId: ev.data.currentListId || null,
          tenantName: ev.data.tenantName || "",
          siteName: ev.data.siteName || ""
        };
      }
    };
    listsLauncherSubsitesListener = (ev) => {
      if (!ev.data || ev.data.__spcsv !== true || ev.data.type !== "SPCSVSubsitesResult") return;
      window.removeEventListener("message", listsLauncherSubsitesListener);
      listsLauncherSubsitesListener = null;
      const subsites = Array.isArray(ev.data.subsites) ? ev.data.subsites : [];
      const subsCacheKey = (panel.dataset.compassSiteKey || getCompassSiteKeyFromLocation() || "").replace(/\/$/, "");
      if (subsCacheKey) {
        siteSubsitesCacheBySite[subsCacheKey] = {
          __spcsv: true,
          type: "SPCSVSubsitesResult",
          subsites: subsites
        };
      }
      if (subsites.length === 0) return;
      applySiteContentsSubsites(panel, subsites);
      if (typeof window.SPOT_noteCompassPanelContentChanged === "function") {
        window.SPOT_noteCompassPanelContentChanged(panel);
      }
    };
    window.addEventListener("message", listsLauncherListsListener);
    window.addEventListener("message", listsLauncherSubsitesListener);
    const listsCacheKey = compassSiteKey.replace(/\/$/, "");
    const cachedListsPayload = listsCacheKey ? siteListsCacheBySite[listsCacheKey] : null;
    const cachedSubsitesPayload = listsCacheKey ? siteSubsitesCacheBySite[listsCacheKey] : null;
    if (cachedListsPayload && listsLauncherListsListener) {
      listsLauncherListsListener({ data: cachedListsPayload });
      if (cachedSubsitesPayload && listsLauncherSubsitesListener) {
        listsLauncherSubsitesListener({ data: cachedSubsitesPayload });
      }
    } else {
    const script = document.createElement("script");
    script.src = chrome.runtime.getURL("getSiteLists.js");
    script.onload = () => script.remove();
    script.onerror = () => {
      if (listsLauncherListsListener) {
        window.removeEventListener("message", listsLauncherListsListener);
        listsLauncherListsListener = null;
      }
      if (listsLauncherSubsitesListener) {
        window.removeEventListener("message", listsLauncherSubsitesListener);
        listsLauncherSubsitesListener = null;
      }
      const loadingEl = panel.querySelector(".sp-toolkit-lists-loading");
      const errorEl = panel.querySelector(".sp-toolkit-lists-error");
      if (loadingEl) loadingEl.style.display = "none";
      if (errorEl) {
        errorEl.textContent = "Failed to load lists.";
        errorEl.style.display = "block";
      }
    };
    (document.head || document.documentElement).appendChild(script);
    }
  }
}

function closeListsLauncherPanel(e) {
  const container = document.getElementById(LISTS_LAUNCHER_ID);
  const panel = document.getElementById(LISTS_LAUNCHER_PANEL_ID);
  if (!container || !panel) return;
  if (e && e.target && container.contains(e.target)) return;
  if (!listsLauncherExpanded) return;
  listsLauncherExpanded = false;
  playLauncherTapeSpin(container, false);
  panel.classList.remove("sp-toolkit-lists-panel-open");
  panel.classList.add("sp-toolkit-panel-rolling-up");
  setTimeout(function () {
    panel.style.display = "none";
    panel.classList.remove("sp-toolkit-panel-rolling-up");
  }, 100);
}

function destroyListsLauncher() {
  const el = document.getElementById(LISTS_LAUNCHER_ID);
  if (el && el.parentNode) el.parentNode.removeChild(el);
}

function initListsLauncher() {
  if (window !== window.top) return;
  refreshCompassShortcutState();
  bindCompassShortcutMacros();
  destroyListsLauncher();
  chrome.storage.local.get("listsLauncherEnabled", (r) => {
    if (r.listsLauncherEnabled === false) return;
    if (document.getElementById(LISTS_LAUNCHER_ID)) return;
    initListsLauncherUI();
  });
}

function initListsLauncherUI() {
  if (window !== window.top) return;
  if (document.getElementById(LISTS_LAUNCHER_ID)) return;
  if (!document.getElementById(LISTS_LAUNCHER_STYLE_ID)) {
    const style = document.createElement("style");
    style.id = LISTS_LAUNCHER_STYLE_ID;
    style.textContent = LISTS_LAUNCHER_CSS + "\n" + LISTS_LAUNCHER_PARITY_CSS;
    document.head.appendChild(style);
  }
  const container = document.createElement("div");
  container.id = LISTS_LAUNCHER_ID;
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "sp-toolkit-lists-btn";
  btn.setAttribute("aria-label", "Site Contents");
  const compassImg = document.createElement("img");
  compassImg.className = "sp-toolkit-launcher-compass";
  compassImg.src = chrome.runtime.getURL("icon.png");
  compassImg.alt = "";
  compassImg.width = 24;
  compassImg.height = 24;
  compassImg.draggable = false;
  btn.appendChild(compassImg);
  btn.addEventListener(
    "dragstart",
    function (e) {
      e.preventDefault();
    },
    true
  );
  const panel = document.createElement("div");
  panel.id = LISTS_LAUNCHER_PANEL_ID;
  panel.className = "sp-toolkit-lists-panel";
  panel.style.display = "none";
  panel.innerHTML =
    "<div class=\"sp-toolkit-lists-panel-header\"><div class=\"sp-toolkit-header-left\">" +
    "<div class=\"sp-toolkit-header-title-row\">" +
    "<div class=\"sp-toolkit-panel-title-label\" aria-live=\"polite\"></div>" +
    "<div class=\"sp-toolkit-header-actions\">" +
    "<button type=\"button\" class=\"sp-toolkit-toolbar-icon-btn sp-toolkit-compass-universal-search-btn\" id=\"btnCompassUniversalSearch\" title=\"Find anything (tabs, links, page info…)\"" +
    " aria-label=\"Open find anything search\">" +
    "<svg class=\"sp-toolkit-compass-header-search-svg\" viewBox=\"0 0 24 24\" aria-hidden=\"true\" focusable=\"false\">" +
    "<path fill=\"currentColor\" fill-rule=\"evenodd\" d=\"M10.5 3.75a6.75 6.75 0 100 13.5 6.75 6.75 0 000-13.5zM2.25 10.5a8.25 8.25 0 1114.59 5.31l3.94 3.94a.75.75 0 11-1.06 1.06l-3.94-3.94A8.25 8.25 0 012.25 10.5z\" clip-rule=\"evenodd\"/>" +
    "</svg></button>" +
    "<button type=\"button\" class=\"sp-toolkit-toolbar-icon-btn sp-toolkit-compass-settings-btn\" id=\"btnCompassPanelSettings\" title=\"Extension settings\" aria-label=\"Open extension settings\">" +
    "<svg class=\"sp-toolkit-compass-header-search-svg\" viewBox=\"0 0 24 24\" aria-hidden=\"true\" focusable=\"false\"><path fill=\"currentColor\" d=\"M11.983 3.5a1 1 0 0 1 .894.553l.535 1.06a6.96 6.96 0 0 1 1.221.508l1.107-.366a1 1 0 0 1 1.09.268l1.061 1.06a1 1 0 0 1 .268 1.09l-.366 1.108c.18.397.34.805.478 1.22l1.09.55a1 1 0 0 1 .552.894v1.5a1 1 0 0 1-.553.894l-1.06.535a6.93 6.93 0 0 1-.508 1.221l.366 1.107a1 1 0 0 1-.268 1.09l-1.06 1.061a1 1 0 0 1-1.09.268l-1.108-.366a6.955 6.955 0 0 1-1.22.478l-.55 1.09a1 1 0 0 1-.894.552h-1.5a1 1 0 0 1-.894-.553l-.535-1.06a6.93 6.93 0 0 1-1.221-.508l-1.107.366a1 1 0 0 1-1.09-.268L5.13 18.74a1 1 0 0 1-.268-1.09l.366-1.108a6.95 6.95 0 0 1-.478-1.22l-1.09-.55a1 1 0 0 1-.552-.894v-1.5a1 1 0 0 1 .553-.894l1.06-.535a6.955 6.955 0 0 1 .508-1.221l-.366-1.107a1 1 0 0 1 .268-1.09L6.19 5.38a1 1 0 0 1 1.09-.268l1.108.366c.397-.18.805-.34 1.22-.478l.55-1.09a1 1 0 0 1 .894-.552h1.5zm.017 5.25a3.5 3.5 0 1 0 0 7.001 3.5 3.5 0 0 0 0-7.001z\"/></svg>" +
    "</button></div>" +
    "<div class=\"dark-toggle\"><button type=\"button\" id=\"spToolkitPanelThemeToggle\" class=\"dark-toggle-btn\" title=\"Early Riser\" aria-label=\"Theme: Early Riser. Click to switch to Night Owl.\"><span class=\"dark-toggle-knob\" aria-hidden=\"true\"></span></button></div>" +
    "<button type=\"button\" class=\"sp-toolkit-toolbar-icon-btn sp-toolkit-compass-close-btn\" id=\"btnCompassPanelClose\" title=\"Close\" aria-label=\"Close panel\">×</button>" +
    "</div>" +
    "<div class=\"sp-toolkit-compass-nav-tabs tab-bar\"></div></div></div>" +
    "<div class=\"sp-toolkit-lists-panel-body\">" +
    "<div class=\"sp-toolkit-compass-panes\">" +
    "<div class=\"sp-toolkit-compass-pane tab-content\" data-compass-pane=\"quicklinks\"><div class=\"sp-toolkit-compass-scroll\"><div class=\"sp-toolkit-compass-quicklinks-host\"></div></div></div>" +
    "<div class=\"sp-toolkit-compass-pane tab-content\" data-compass-pane=\"context\"><div class=\"sp-toolkit-compass-scroll\" id=\"contextPanel\">" +
    "<div class=\"toolbar\">" +
    "<input type=\"text\" id=\"contextFilterInput\" class=\"sp-toolkit-compass-context-filter\" placeholder=\"Filter properties and values…\" />" +
    "<button type=\"button\" id=\"btnRefreshContext\" class=\"sp-toolkit-compass-secondary sp-toolkit-btn-refresh-context btn-ghost btn-ghost-icon\" title=\"Refresh page properties\" aria-label=\"Refresh page properties\">" +
    "<svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" stroke-linecap=\"round\" stroke-linejoin=\"round\" aria-hidden=\"true\"><path d=\"M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8\"/><path d=\"M21 3v5h-5\"/><path d=\"M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16\"/><path d=\"M8 16H3v5\"/></svg>" +
    "</button></div>" +
    "<div id=\"contextCount\" class=\"sp-toolkit-compass-context-count sp-toolkit-compass-muted\">Properties: —</div>" +
    "<div class=\"schema-col-header compass-context-schema-head\"><span>Property</span><span>Value</span></div>" +
    "<div id=\"contextList\" class=\"sp-toolkit-compass-context-host\"></div></div></div>" +
    "<div class=\"sp-toolkit-compass-pane tab-content\" data-compass-pane=\"columns\"><div class=\"sp-toolkit-compass-scroll\" id=\"searchSchemaPanel\">" +
    "<div class=\"toolbar\">" +
    "<input type=\"text\" id=\"searchSchemaFilter\" class=\"sp-toolkit-compass-columns-filter\" placeholder=\"Search columns…\" />" +
    "<button type=\"button\" id=\"btnOpenColumnCreator\" class=\"sp-toolkit-compass-secondary sp-toolkit-btn-open-column-creator btn-ghost btn-ghost-icon\" title=\"Create column\" aria-label=\"Create column\" aria-expanded=\"false\">" +
    "<svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" stroke-linecap=\"round\" aria-hidden=\"true\"><path d=\"M12 5v14M5 12h14\"/></svg>" +
    "</button></div>" +
    "<div class=\"sp-toolkit-compass-column-creator-host\"></div>" +
    "<div id=\"searchSchemaCount\" class=\"sp-toolkit-compass-columns-count sp-toolkit-compass-muted\">Columns: —</div>" +
    "<div id=\"searchSchemaError\" class=\"sp-toolkit-compass-columns-error sp-toolkit-compass-err\" style=\"display:none\"></div>" +
    "<div class=\"schema-col-header\"><span>Display Name</span><span>Internal Name</span><span>Crawled Property</span><span>Type</span></div>" +
    "<div id=\"searchSchemaList\" class=\"sp-toolkit-compass-columns-list\"></div></div></div>" +
    "<div class=\"sp-toolkit-compass-pane tab-content\" data-compass-pane=\"reports\"><div class=\"sp-toolkit-compass-scroll\"><div class=\"sp-toolkit-compass-reports-host\"></div></div></div>" +
    "<div class=\"sp-toolkit-compass-pane tab-content\" data-compass-pane=\"refinableProps\"><div class=\"sp-toolkit-compass-scroll\" id=\"refinablePropsPanel\">" +
    "<div class=\"sp-toolkit-compass-field\"><label class=\"sp-toolkit-compass-label\" for=\"refinableTypeSelect\">Refinable type</label><select id=\"refinableTypeSelect\" class=\"sp-toolkit-compass-refinable-type\"></select></div>" +
    "<div class=\"sp-toolkit-compass-field\"><label class=\"sp-toolkit-compass-label\" for=\"refinableNumberSelect\">Number</label><select id=\"refinableNumberSelect\" class=\"sp-toolkit-compass-refinable-num\"></select></div>" +
    "<p class=\"sp-toolkit-compass-muted\">Opens the search schema managed property page for the selected refinable (e.g. RefinableString00).</p>" +
    "<div id=\"refinableMappingsOut\" class=\"sp-toolkit-compass-refinable-out sp-toolkit-compass-muted\">Select a refinable, then mappings load when you open this tab.</div>" +
    "<button type=\"button\" id=\"btnRefinableConfigure\" class=\"sp-toolkit-compass-primary sp-toolkit-btn-refinable-config btn-configure\"><span class=\"sp-toolkit-btn-gear\" aria-hidden=\"true\">⚙</span><span>Configure</span></button></div></div>" +
    "<div class=\"sp-toolkit-compass-pane tab-content\" data-compass-pane=\"viewManager\"><div class=\"sp-toolkit-compass-scroll\"><div class=\"sp-toolkit-compass-views-host\"></div></div></div>" +
    "<div class=\"sp-toolkit-compass-pane tab-content\" data-compass-pane=\"toolkitSettings\"><div class=\"sp-toolkit-compass-scroll\"><div class=\"sp-toolkit-compass-settings-host\"></div></div></div>" +
    "<div class=\"sp-toolkit-compass-pane tab-content sp-toolkit-compass-pane-active active\" data-compass-pane=\"siteContents\">" +
    "<div class=\"sp-toolkit-site-contents-stack\">" +
    "<div class=\"sp-toolkit-contents-wrap\">" +
    "<div class=\"sp-toolkit-lists-loading\"><span class=\"sp-toolkit-lists-loading-dot\"></span><span class=\"sp-toolkit-lists-loading-dot\"></span><span class=\"sp-toolkit-lists-loading-dot\"></span><span>Loading…</span></div>" +
    "<div class=\"sp-toolkit-lists-error\" style=\"display:none\"></div>" +
    "<table class=\"sp-toolkit-contents-table\"><thead></thead><tbody></tbody></table>" +
    "</div></div></div></div>";
  panel.addEventListener("click", function panelInnerClick(e) {
    e.stopPropagation();
  });
  const themeToggleBtn = panel.querySelector("#spToolkitPanelThemeToggle");
  const settingsBtn = panel.querySelector("#btnCompassPanelSettings");
  const closeBtn = panel.querySelector("#btnCompassPanelClose");
  if (themeToggleBtn) {
    themeToggleBtn.addEventListener("click", function (e) {
      e.preventDefault();
      e.stopPropagation();
      const nextDark = !panel.classList.contains("sp-toolkit-lists-panel-dark");
      setListsPanelThemeClasses(panel, nextDark);
      chrome.storage.local.set({ darkMode: nextDark });
    });
  }
  if (settingsBtn) {
    settingsBtn.addEventListener("click", function (e) {
      e.preventDefault();
      e.stopPropagation();
      if (typeof window.SPOT_openCompassToolkitSettings === "function") {
        window.SPOT_openCompassToolkitSettings();
      }
    });
  }
  if (closeBtn) {
    closeBtn.addEventListener("click", function (e) {
      e.preventDefault();
      e.stopPropagation();
      if (listsLauncherExpanded) toggleListsLauncher();
    });
  }
  container.appendChild(panel);
  container.appendChild(btn);
  document.body.appendChild(container);
  chrome.storage.local.get([LISTS_LAUNCHER_POSITION_KEY, LISTS_LAUNCHER_ICON_SCALE_KEY], function (r) {
    applyLauncherIconScale(container, r[LISTS_LAUNCHER_ICON_SCALE_KEY]);
    applyLauncherPositionFromStored(container, r[LISTS_LAUNCHER_POSITION_KEY]);
  });
  attachListsLauncherDragAndTap(container, btn);
  bindListsLauncherResizeClamp(container);
  document.addEventListener("click", closeListsLauncherPanel);
  applyListsPanelTheme(panel);
  chrome.storage.local.get(LISTS_LAUNCHER_ANIMATION_SECONDS_KEY, function (r) {
    applyListsPanelAnimationSpeed(panel, r[LISTS_LAUNCHER_ANIMATION_SECONDS_KEY]);
  });
}

function setListsPanelThemeClasses(panel, isDark) {
  if (!panel) return;
  const container = panel.parentNode && panel.parentNode.id === LISTS_LAUNCHER_ID ? panel.parentNode : document.getElementById(LISTS_LAUNCHER_ID);
  panel.classList.remove("sp-toolkit-lists-panel-dark", "sp-toolkit-lists-panel-light");
  panel.classList.add(isDark ? "sp-toolkit-lists-panel-dark" : "sp-toolkit-lists-panel-light");
  if (container) {
    container.classList.remove("sp-toolkit-launcher-dark", "sp-toolkit-launcher-light");
    container.classList.add(isDark ? "sp-toolkit-launcher-dark" : "sp-toolkit-launcher-light");
  }
  const themeToggleBtn = panel.querySelector("#spToolkitPanelThemeToggle");
  if (themeToggleBtn) {
    themeToggleBtn.title = isDark ? "Night Owl" : "Early Riser";
    themeToggleBtn.setAttribute("aria-label", isDark ? "Theme: Night Owl. Click to switch to Early Riser." : "Theme: Early Riser. Click to switch to Night Owl.");
  }
}

function applyListsPanelTheme(panel, opts) {
  if (!panel) return;
  const fromStorage = !!(opts && opts.fromStorage);
  chrome.storage.local.get("darkMode", (r) => {
    const isDark = !!r.darkMode;
    const alreadySet = panel.classList.contains(isDark ? "sp-toolkit-lists-panel-dark" : "sp-toolkit-lists-panel-light");
    if (alreadySet) {
      setListsPanelThemeClasses(panel, isDark);
      return;
    }
    if (!fromStorage) panel.classList.add("sp-toolkit-theme-switching");
    setListsPanelThemeClasses(panel, isDark);
    if (!fromStorage) {
      requestAnimationFrame(function () {
        panel.classList.remove("sp-toolkit-theme-switching");
      });
    }
  });
}

chrome.storage.onChanged.addListener((changes, areaName) => {
  if (window !== window.top || areaName !== "local") return;
  if (changes.listsLauncherEnabled) {
    const enabled = changes.listsLauncherEnabled.newValue !== false;
    if (enabled) requestAnimationFrame(function () { requestAnimationFrame(function () { initListsLauncherUI(); }); });
    else destroyListsLauncher();
    return;
  }
  if (changes.darkMode) {
    const panel = document.getElementById(LISTS_LAUNCHER_PANEL_ID);
    if (panel) applyListsPanelTheme(panel, { fromStorage: true });
  }
  if (changes[COMPASS_SHORTCUTS_ENABLED_KEY] || changes[COMPASS_SHORTCUTS_KEY]) {
    refreshCompassShortcutState();
  }
  if (changes[LISTS_LAUNCHER_ANIMATION_SECONDS_KEY]) {
    const panel = document.getElementById(LISTS_LAUNCHER_PANEL_ID);
    if (panel) applyListsPanelAnimationSpeed(panel, changes[LISTS_LAUNCHER_ANIMATION_SECONDS_KEY].newValue);
  }
  if (changes[LISTS_LAUNCHER_POSITION_KEY]) {
    const c = document.getElementById(LISTS_LAUNCHER_ID);
    if (c) applyLauncherPositionFromStored(c, changes[LISTS_LAUNCHER_POSITION_KEY].newValue);
  }
  if (changes[LISTS_LAUNCHER_ICON_SCALE_KEY]) {
    const c = document.getElementById(LISTS_LAUNCHER_ID);
    if (c) applyLauncherIconScale(c, changes[LISTS_LAUNCHER_ICON_SCALE_KEY].newValue);
  }
});

function setupListeners() {
  window.removeEventListener("message", onPageMessage);
  window.addEventListener("message", onPageMessage);
}

let matrixWorkerUnloadHooked = false;
function enableMatrixWorkerUnloadWarning() {
  if (matrixWorkerUnloadHooked) return;
  matrixWorkerUnloadHooked = true;
  window.addEventListener("beforeunload", function (e) {
    if (window.__SPOToolkitExportRunning) {
      e.preventDefault();
      e.returnValue = "Permissions matrix export is still running. Closing this tab will stop the export.";
      return e.returnValue;
    }
  });
}

function cleanExportHeadline(msg) {
  return String(msg || "").replace(/\s*\(still working…\)+/gi, "").trim();
}

function spotRefreshExportProgressUI() {
  if (typeof window.SPOT_refreshExportProgress === "function") {
    window.SPOT_refreshExportProgress();
  }
}

function exportDownloadTimeoutMs(detail) {
  let bytes = 0;
  if (detail && detail.text != null) bytes = String(detail.text).length;
  else if (detail && detail.bufferBase64) bytes = Math.floor(String(detail.bufferBase64).length * 0.75);
  else if (detail && detail.buffer) bytes = detail.buffer.byteLength || 0;
  return Math.min(900000, Math.max(60000, 60000 + Math.floor(bytes / (512 * 1024)) * 5000));
}

function waitForChromeDownload(downloadId, objectUrl, timeoutMs) {
  timeoutMs = timeoutMs || exportDownloadTimeoutMs();
  return new Promise((resolve, reject) => {
    let settled = false;
    const finish = (err) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      clearInterval(pollTimer);
      try { chrome.downloads.onChanged.removeListener(onChanged); } catch (_) {}
      if (objectUrl) {
        setTimeout(function () {
          try { URL.revokeObjectURL(objectUrl); } catch (_) {}
        }, 5000);
      }
      if (err) reject(err);
      else resolve();
    };
    const checkItem = (item) => {
      if (!item) return;
      if (item.state === "complete") finish(null);
      else if (item.state === "interrupted") finish(new Error(item.error || "Download interrupted"));
    };
    const onChanged = (delta) => {
      if (delta.id !== downloadId) return;
      if (delta.error && delta.error.current) finish(new Error(String(delta.error.current)));
      if (!delta.state || !delta.state.current) return;
      if (delta.state.current === "complete") finish(null);
      else if (delta.state.current === "interrupted") finish(new Error("Download interrupted"));
    };
    const timer = setTimeout(function () { finish(new Error("Download timed out")); }, timeoutMs);
    const pollTimer = setInterval(function () {
      if (settled) return;
      try {
        chrome.downloads.search({ id: downloadId }, (items) => {
          checkItem(items && items[0]);
        });
      } catch (_) {}
    }, 2000);
    try { chrome.downloads.onChanged.addListener(onChanged); } catch (_) {}
    try {
      chrome.downloads.search({ id: downloadId }, (items) => {
        checkItem(items && items[0]);
      });
    } catch (_) {}
  });
}

function normalizeExportDownloadDetail(detail) {
  detail = detail || {};
  if (detail.bufferBase64) {
    const binary = atob(String(detail.bufferBase64));
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    return {
      path: detail.path,
      mime: detail.mime,
      buffer: bytes.buffer
    };
  }
  return detail;
}

async function performExportDownload(detail, options) {
  options = options || {};
  const waitForComplete = options.waitForComplete !== false;
  detail = normalizeExportDownloadDetail(detail);
  if (!chrome.downloads || typeof chrome.downloads.download !== "function") {
    throw new Error("Downloads API unavailable — reload the extension.");
  }
  let blob;
  if (detail.text != null) {
    blob = new Blob([detail.text], { type: detail.mime || "text/csv;charset=utf-8" });
  } else if (detail.buffer) {
    blob = new Blob([detail.buffer], { type: detail.mime || "application/octet-stream" });
  } else {
    throw new Error("Empty download payload");
  }
  const objectUrl = URL.createObjectURL(blob);
  const filename = String(detail.path || "export.dat").replace(/\\/g, "/");
  const timeoutMs = exportDownloadTimeoutMs(detail);
  await new Promise((resolve, reject) => {
    chrome.downloads.download({
      url: objectUrl,
      filename,
      conflictAction: "uniquify",
      saveAs: false
    }, (downloadId) => {
      if (chrome.runtime.lastError || downloadId == null) {
        URL.revokeObjectURL(objectUrl);
        reject(new Error(chrome.runtime.lastError?.message || "Download failed"));
        return;
      }
      if (!waitForComplete) {
        resolve();
        waitForChromeDownload(downloadId, objectUrl, timeoutMs).catch(function () {
          setTimeout(function () {
            try { URL.revokeObjectURL(objectUrl); } catch (_) {}
          }, 15000);
        });
        return;
      }
      waitForChromeDownload(downloadId, objectUrl, timeoutMs).then(resolve, reject);
    });
  });
}

const exportDownloadChunkTransfers = new Map();

function mergeExportDownloadChunks(state) {
  const merged = new Uint8Array(state.totalBytes);
  let offset = 0;
  for (let i = 0; i < state.totalChunks; i++) {
    const chunk = state.chunks[i];
    if (!chunk) throw new Error("Missing download chunk " + (i + 1) + " of " + state.totalChunks);
    const view = chunk instanceof ArrayBuffer
      ? new Uint8Array(chunk)
      : new Uint8Array(chunk.buffer, chunk.byteOffset, chunk.byteLength);
    merged.set(view, offset);
    offset += view.byteLength;
  }
  return merged.buffer;
}

function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  const chunk = 8192;
  const parts = [];
  for (let i = 0; i < bytes.length; i += chunk) {
    parts.push(String.fromCharCode.apply(null, bytes.subarray(i, i + chunk)));
  }
  return btoa(parts.join(""));
}

async function relayExportDownloadPerform(detail) {
  const normalized = normalizeExportDownloadDetail(detail || {});
  if (!normalized.path && detail && detail.path) normalized.path = detail.path;
  if (normalized.text == null && !normalized.buffer) {
    throw new Error("Export download request missing payload.");
  }
  const payload = {
    path: normalized.path,
    mime: normalized.mime
  };
  if (normalized.text != null) payload.text = normalized.text;
  else payload.bufferBase64 = arrayBufferToBase64(normalized.buffer);
  const timeoutMs = Math.min(45000, exportDownloadTimeoutMs(normalized));
  const resp = await new Promise((resolve, reject) => {
    let settled = false;
    const timer = setTimeout(function () {
      if (settled) return;
      settled = true;
      reject(new Error("Download relay timed out"));
    }, timeoutMs);
    try {
      chrome.runtime.sendMessage({ type: "SPCSVExportDownloadPerform", detail: payload }, (response) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        if (chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message || "Download relay failed"));
          return;
        }
        resolve(response || { ok: false, error: "Download handler did not respond" });
      });
    } catch (err) {
      clearTimeout(timer);
      reject(err);
    }
  });
  if (!resp || !resp.ok) {
    throw new Error((resp && resp.error) || "Download failed");
  }
}

async function relayExportDownloadRequest(detail) {
  const requestId = detail && detail.requestId;
  const reply = (ok, error) => {
    window.postMessage({
      __spcsv: true,
      type: "SPCSVExportDownloadResult",
      detail: { requestId, ok: !!ok, error: error || "" }
    }, "*");
  };
  if (!requestId) return;
  try {
    await relayExportDownloadPerform(detail);
    reply(true);
  } catch (err) {
    reply(false, err && err.message ? err.message : String(err));
  }
}

async function handleExportDownloadChunk(detail) {
  detail = detail || {};
  const requestId = detail.requestId;
  if (!requestId) return;
  let state = exportDownloadChunkTransfers.get(requestId);
  if (!state) {
    state = {
      requestId,
      path: detail.path,
      mime: detail.mime,
      totalBytes: detail.totalBytes || 0,
      totalChunks: detail.totalChunks || 0,
      chunks: []
    };
    exportDownloadChunkTransfers.set(requestId, state);
  }
  state.chunks[detail.chunkIndex] = detail.buffer;
}

async function finalizeExportDownloadTransfer(detail) {
  detail = detail || {};
  const requestId = detail.requestId;
  if (!requestId) return;
  const state = exportDownloadChunkTransfers.get(requestId);
  if (!state) {
    window.postMessage({
      __spcsv: true,
      type: "SPCSVExportDownloadResult",
      detail: { requestId, ok: false, error: "Download chunks missing" }
    }, "*");
    return;
  }
  exportDownloadChunkTransfers.delete(requestId);
  try {
    const buffer = mergeExportDownloadChunks(state);
    await relayExportDownloadPerform({
      requestId,
      path: state.path,
      mime: state.mime,
      buffer
    });
    window.postMessage({
      __spcsv: true,
      type: "SPCSVExportDownloadResult",
      detail: { requestId, ok: true, error: "" }
    }, "*");
  } catch (err) {
    window.postMessage({
      __spcsv: true,
      type: "SPCSVExportDownloadResult",
      detail: {
        requestId,
        ok: false,
        error: err && err.message ? err.message : String(err)
      }
    }, "*");
  }
}

function onPageMessage(e) {
  if (!e.data || e.data.__spcsv !== true) return;
  const d = e.data.detail || {};
  if (e.data.type === "SPCSVExportStarted") {
    const ribbon = document.getElementById(PROGRESS_BOX_ID);
    if (ribbon && ribbon.parentNode) ribbon.parentNode.removeChild(ribbon);
    try {
      chrome.runtime.sendMessage({
        type: "SPCSVExportStarted",
        report: d.report || "export",
        message: d.message || "Export started…"
      });
    } catch (err) {}
    spotRefreshExportProgressUI();
  } else if (e.data.type === "SPCSVExportProgress") {
    const msg = cleanExportHeadline(d.message || "Export in progress…");
    const percent = d.percent != null ? d.percent : undefined;
    const logLine = d.pulse ? null : (d.logLine || d.message || msg);
    if (window.SPOToolkitMatrixWorkerLock && window.SPOToolkitMatrixWorkerLock.isActive()) {
      window.SPOToolkitMatrixWorkerLock.update({ message: msg, percent: percent, logLine: logLine || undefined });
    }
    try {
      chrome.runtime.sendMessage({
        type: "SPCSVExportProgress",
        report: d.report || "export",
        message: msg,
        logLine: logLine || undefined,
        percent,
        pulse: d.pulse || undefined
      });
    } catch (err) {}
    spotRefreshExportProgressUI();
  } else if (e.data.type === "SPCSVExportDone") {
    window.__SPOToolkitExportRunning = false;
    const success = !!(d && d.success);
    const message = (d && d.message) || (success ? "Done!" : "Export failed.");
    const panelEl = document.getElementById(LISTS_LAUNCHER_PANEL_ID);
    const rs = panelEl && panelEl.querySelector(".sp-toolkit-compass-reports-status");
    if (rs) {
      rs.textContent = message;
      rs.className = "sp-toolkit-compass-reports-status" + (success ? " ok" : "");
    }
    try {
      chrome.runtime.sendMessage({ type: "SPCSVExportDone", success, message });
    } catch (err) {}
    spotRefreshExportProgressUI();
  } else if (e.data.type === "SPCSVExportCancel") {
    try {
      chrome.runtime.sendMessage({ type: "SPCSVExportCancel" });
    } catch (err) {}
  } else if (e.data.type === "SPCSVExportDownloadRequest") {
    void relayExportDownloadRequest(d.detail || {});
  } else if (e.data.type === "SPCSVExportDownloadChunk") {
    void handleExportDownloadChunk(d.detail || {});
  } else if (e.data.type === "SPCSVExportDownloadChunkEnd") {
    void finalizeExportDownloadTransfer(d.detail || {});
  }
}

/** JSON script node read by createColumn.js (page context). */
const SP_COLUMN_CREATE_PARAMS_SCRIPT_ID = "sp-column-create-params";

function attachSpColumnCreateParamsScript(message) {
  const payload = {
    siteUrl: message.siteUrl || "",
    listId: message.listId || "",
    target: message.target || "list",
    schemaXml: message.schemaXml || "",
    siteFieldInternal: message.siteFieldInternal || "",
  };
  let el = document.getElementById(SP_COLUMN_CREATE_PARAMS_SCRIPT_ID);
  if (el) el.remove();
  el = document.createElement("script");
  el.id = SP_COLUMN_CREATE_PARAMS_SCRIPT_ID;
  el.type = "application/json";
  el.textContent = JSON.stringify(payload);
  (document.head || document.documentElement).appendChild(el);
}

// Inject a script and wait for one postMessage response. Handles timeout and script load error.
function injectAndWait(scriptName, messageType, parseData, sendResponse, options = {}) {
  const timeoutMs = options.timeoutMs ?? SCRIPT_TIMEOUT;
  const errorPayload = options.errorPayload ?? { ok: false, error: "Timeout or load failed" };
  let done = false;
  function finish(payload) {
    if (done) return;
    done = true;
    clearTimeout(tid);
    window.removeEventListener("message", listener);
    sendResponse(payload);
  }
  const listener = (ev) => {
    if (!ev.data || ev.data.__spcsv !== true || ev.data.type !== messageType) return;
    finish(parseData(ev.data));
  };
  if (options.beforeInject) options.beforeInject();
  const script = document.createElement("script");
  script.src = chrome.runtime.getURL(scriptName) + "?cb=" + Date.now();
  script.onload = () => script.remove();
  script.onerror = () => finish(errorPayload);
  window.addEventListener("message", listener);
  const tid = setTimeout(() => finish(errorPayload), timeoutMs);
  (document.head || document.documentElement).appendChild(script);
}

function runExportInjection(message, sendResponse, options) {
  options = options || {};
  window.__SPOToolkitExportCancel = false;
  window.__SPOToolkitExportRunning = true;
  if (!options.skipStarted) {
    showProgress("Starting export");
    setupListeners();
    try {
      chrome.runtime.sendMessage({
        type: "SPCSVExportStarted",
        report: message.report === "permissionsMatrix" ? "permissionsMatrix" : (message.report || "export"),
        message: message.report === "permissionsMatrix" ? "Permissions matrix export started…" : "Export started…"
      });
    } catch (_) {}
  } else {
    setupListeners();
  }
  const { siteUrl, listId, viewId, exportFilename, pageLimit, includeVersions, selectedColumns, report, format, matrixIncludeSubsites, matrixIncludeAllInherited } = message;
  const isMatrixReport = report === "permissionsMatrix";
  const exportScriptName = isMatrixReport ? "permissionsMatrixExport.js" : "exportCSV.js";
  const params = {
    u: (siteUrl || "").replace(/\/$/, ""),
    lid: listId || "",
    vid: viewId || "",
    f: exportFilename || "",
    pl: pageLimit ?? 1000,
    incVer: includeVersions !== false
  };
  if (report) params.report = report;
  if (format) params.format = format;
  if (selectedColumns && Array.isArray(selectedColumns) && selectedColumns.length > 0) params.cols = selectedColumns;
  if (isMatrixReport) {
    params.matrixIncludeSubsites = matrixIncludeSubsites !== false;
    params.matrixIncludeAllInherited = matrixIncludeAllInherited === true;
    params.matrixExpandGroups = message.matrixExpandGroups === true;
    params.matrixIncludeFolderSharingLinks = message.matrixIncludeFolderSharingLinks !== false;
    if (message.matrixSharingLinkFetchAll === false) params.matrixSharingLinkFetchAll = false;
    if (message.matrixMaxListItems != null) params.matrixMaxListItems = parseInt(message.matrixMaxListItems, 10) || 2000;
    if (message.matrixListItemPageSize != null) params.matrixListItemPageSize = parseInt(message.matrixListItemPageSize, 10) || 5000;
    if (Array.isArray(message.matrixSelectedPaths) && message.matrixSelectedPaths.length) {
      params.matrixSelectedPaths = message.matrixSelectedPaths;
    }
  }
  if (!isMatrixReport && Array.isArray(message.reportSelectedLists) && message.reportSelectedLists.length) {
    params.reportSelectedLists = message.reportSelectedLists;
    if (message.rootSiteTitle) params.rootSiteTitle = message.rootSiteTitle;
  }
  let el = document.getElementById("spcsv-params-json");
  if (el) el.remove();
  el = document.createElement("script");
  el.id = "spcsv-params-json";
  el.type = "application/json";
  el.textContent = JSON.stringify(params);
  (document.head || document.documentElement).appendChild(el);

  function injectExportScript() {
    const script = document.createElement("script");
    script.src = chrome.runtime.getURL(exportScriptName);
    script.onload = () => { script.remove(); sendResponse({ ok: true, message: "Export started. Watch the progress bar for status." }); };
    script.onerror = () => {
      finishProgress(false, "Failed to load export script.");
      sendResponse({ ok: false, error: "Failed to load " + exportScriptName });
    };
    (document.head || document.documentElement).appendChild(script);
  }

  function loadScriptChain(urls, onDone, onErr) {
    if (!urls.length) { onDone(); return; }
    const s = document.createElement("script");
    s.src = urls[0];
    s.onload = () => { s.remove(); loadScriptChain(urls.slice(1), onDone, onErr); };
    s.onerror = onErr;
    (document.head || document.documentElement).appendChild(s);
  }

  const folderBundleReport = !isMatrixReport && (report === "exportCSV" || report === "folderCount" || report === "pathLengths");
  const needsXlsxPreload = format === "xlsx" || isMatrixReport || folderBundleReport;
  if (needsXlsxPreload) {
    const scriptChain = [chrome.runtime.getURL("jszip.min.js")];
    if (isMatrixReport) scriptChain.push(chrome.runtime.getURL("permissionsMatrixStyles.js"));
    scriptChain.push(chrome.runtime.getURL(exportScriptName));
    loadScriptChain(scriptChain, () => {
      sendResponse({ ok: true, message: "Export started. Watch the progress bar for status." });
    }, () => {
      finishProgress(false, "Failed to load export script.");
      sendResponse({ ok: false, error: "Failed to load JSZip or export script" });
    });
  } else {
    injectExportScript();
  }
}

function tryResumeMatrixExportInTab(prefetched) {
  if (!/sharepoint\.com/i.test(location.href || "")) return;
  function attemptResume(data) {
    if (!data || !data.active || data.report !== "permissionsMatrix" || !data.exportParams) return;
    if (window.__SPOToolkitExportRunning || window.__SPOToolkitMatrixResumeAttempted) return;
    window.__SPOToolkitMatrixResumeAttempted = true;
    chrome.tabs.getCurrent(function (tab) {
      if (chrome.runtime.lastError || !tab || tab.id == null) {
        window.__SPOToolkitMatrixResumeAttempted = false;
        return;
      }
      if (data.workerTabId != null && tab.id !== data.workerTabId) {
        window.__SPOToolkitMatrixResumeAttempted = false;
        return;
      }
      if (window.SPOToolkitMatrixWorkerLock) window.SPOToolkitMatrixWorkerLock.show();
      enableMatrixWorkerUnloadWarning();
      setupListeners();
      try {
        chrome.runtime.sendMessage({
          type: "SPCSVExportProgress",
          report: "permissionsMatrix",
          message: "Resuming permissions matrix export after refresh…",
          logLine: "Resuming permissions matrix export after refresh…",
          percent: data.percent
        });
      } catch (_) {}
      runExportInjection(data.exportParams, function () {}, { skipStarted: true });
    });
  }
  if (prefetched) {
    attemptResume(prefetched);
    return;
  }
  try {
    chrome.storage.session.get("spcsvExportProgress", function (result) {
      if (chrome.runtime.lastError || !result) return;
      attemptResume(result.spcsvExportProgress);
    });
  } catch (_) {}
}

function dispatchToolkitMessage(message, sendResponse) {
  if (message.type === "SPCSVExportPing") {
    sendResponse({ running: window.__SPOToolkitExportRunning === true });
    return false;
  }
  if (message.action === "performExportDownload") {
    performExportDownload(message.detail || {}, { waitForComplete: message.waitForComplete !== false })
      .then(function () { sendResponse({ ok: true }); })
      .catch(function (err) {
        sendResponse({ ok: false, error: err && err.message ? err.message : String(err) });
      });
    return true;
  }
  if (message.action === "cancelExportCSV") {
    window.__SPOToolkitExportRunning = false;
    window.__SPOToolkitExportCancel = true;
    try {
      const s = document.createElement("script");
      s.textContent = "window.__SPOToolkitExportCancel=true;try{window.dispatchEvent(new CustomEvent('spotoolkit-export-cancel'));}catch(e){}";
      (document.documentElement || document.head).appendChild(s);
      s.remove();
    } catch (_) {}
    try { window.postMessage({ __spcsv: true, type: "SPCSVExportCancel" }, "*"); } catch (_) {}
    sendResponse({ ok: true });
    return false;
  }
  if (message.action === "matrixWorkerLockAndRun") {
    if (window.SPOToolkitMatrixWorkerLock) {
      const report = (message.exportMessage && message.exportMessage.report) || "exportCSV";
      window.SPOToolkitMatrixWorkerLock.show(report);
    }
    enableMatrixWorkerUnloadWarning();
    runExportInjection(message.exportMessage, sendResponse, { skipStarted: true });
    return true;
  }
  if (message.action === "matrixWorkerFinish") {
    if (window.SPOToolkitMatrixWorkerLock) {
      window.SPOToolkitMatrixWorkerLock.finish(!!message.success, message.message || "", message.autoCloseMs || 5000);
    }
    window.__SPOToolkitExportRunning = false;
    sendResponse({ ok: true });
    return false;
  }
  if (message.action === "runExportCSVWorker" && message._workerFrame) {
    if (window.SPOToolkitMatrixWorkerLock) window.SPOToolkitMatrixWorkerLock.show();
    enableMatrixWorkerUnloadWarning();
    runExportInjection(message, sendResponse, { skipStarted: true });
    return true;
  }
  if (message.type === "SPOToolkitPopupOpened") {
    const launcher = document.getElementById(LISTS_LAUNCHER_ID);
    const panel = document.getElementById(LISTS_LAUNCHER_PANEL_ID);
    if (launcher) launcher.classList.remove("sp-toolkit-launcher-rolled-up");
    if (panel && listsLauncherExpanded) {
      listsLauncherExpanded = false;
      panel.classList.remove("sp-toolkit-lists-panel-open");
      panel.classList.add("sp-toolkit-panel-rolling-up");
      setTimeout(function () {
        panel.style.display = "none";
        panel.classList.remove("sp-toolkit-panel-rolling-up");
      }, 130);
    }
    sendResponse({});
    return true;
  }
  if (message.type === "SPOToolkitPopupClosed") {
    const launcher = document.getElementById(LISTS_LAUNCHER_ID);
    if (launcher) launcher.classList.remove("sp-toolkit-launcher-rolled-up");
    sendResponse({});
    return true;
  }
  if (message.type === "SPOToolkitOpenSiteContentsPanel") {
    chrome.storage.local.get("listsLauncherEnabled", (r) => {
      if (r.listsLauncherEnabled === false) {
        sendResponse({ ok: false, error: "disabled" });
        return;
      }
      try {
        if (!document.getElementById(LISTS_LAUNCHER_ID)) {
          initListsLauncherUI();
        }
        const container = document.getElementById(LISTS_LAUNCHER_ID);
        const panel = document.getElementById(LISTS_LAUNCHER_PANEL_ID);
        if (!container || !panel) {
          sendResponse({ ok: false, error: "not_initialized" });
          return;
        }
        if (!listsLauncherExpanded) {
          toggleListsLauncher();
        }
        sendResponse({ ok: true });
      } catch (_) {
        sendResponse({ ok: false, error: "exception" });
      }
    });
    return true;
  }
  if (message.action === "showFact") {
    showFactToast(message.fact || "", message.sourceUrl || "");
    sendResponse({});
    return true;
  }
  if (message.action === "getPageContext") {
    injectAndWait(
      "getPageContext.js",
      "SPCSVPageContext",
      (data) => {
        const web = (data.webAbsoluteUrl || "").replace(/\/$/, "");
        const site = (data.siteAbsoluteUrl || "").replace(/\/$/, "");
        const primaryWeb = web || site;
        if (!primaryWeb) return { ok: false, error: "No context from page" };
        return {
          ok: true,
          webAbsoluteUrl: primaryWeb,
          siteAbsoluteUrl: site,
          pageListId: data.pageListId || "",
          listUrl: data.listUrl || "",
          siteId: data.siteId || "",
          webId: data.webId || ""
        };
      },
      sendResponse,
      { timeoutMs: 5000, errorPayload: { ok: false, error: "No context from page" } }
    );
    return true;
  }

  if (message.action === "getPageContextFull") {
    if (pageContextFullCacheHref === location.href && pageContextFullCacheData) {
      sendResponse({
        ok: true,
        data: pageContextFullCacheData,
        error: null
      });
      return true;
    }
    injectAndWait(
      "getPageContextFull.js",
      "SPCSVPageContextFull",
      (data) => {
        const result = {
          ok: !!(data && data.data && typeof data.data === "object"),
          data: data && data.data && typeof data.data === "object" ? data.data : {},
          error: data && data.data ? null : "No full page context from page"
        };
        if (result.ok && Object.keys(result.data).length) {
          pageContextFullCacheHref = location.href;
          pageContextFullCacheData = result.data;
        }
        return result;
      },
      sendResponse,
      { timeoutMs: 6000, errorPayload: { ok: false, error: "No full page context from page", data: {} } }
    );
    return true;
  }

  if (message.action === "getViewFormatContext") {
    injectAndWait(
      "getViewFormatContext.js",
      "SPCSVViewFormatContext",
      (data) => {
        const normGuid = (s) =>
          String(s || "")
            .replace(/[{}]/g, "")
            .replace(/%7B|%7D/gi, "")
            .trim();
        const listId = normGuid(data.listId);
        const viewId = normGuid(data.viewId);
        const webAbsoluteUrl = String(data.webAbsoluteUrl || "").replace(/\/$/, "");
        if (!listId || !webAbsoluteUrl) {
          return {
            ok: false,
            error: "Could not read list/site. Open a list or library view on SharePoint.",
          };
        }
        let sitePath = "/";
        try {
          sitePath = new URL(webAbsoluteUrl).pathname.replace(/\/$/, "") || "/";
        } catch (_) {}
        return { ok: true, listId, viewId, sitePath };
      },
      sendResponse,
      { timeoutMs: 6000, errorPayload: { ok: false, error: "Timeout reading page context." } }
    );
    return true;
  }

  if (message.action === "getPageContextJson") {
    injectAndWait(
      "getPageContextJson.js",
      "SPCSVPageContextJson",
      (data) => (data.error ? { ok: false, error: data.error } : { ok: true, data: data.data || {} }),
      sendResponse,
      { timeoutMs: 12000, errorPayload: { ok: false, error: "Timeout or page did not return context" } }
    );
    return true;
  }

  if (message.action === "rest") {
    const method = (message.method || "GET").toUpperCase();
    const path = message.path || "";
    const body = message.body;
    (async () => {
      try {
        let digest = null;
        if (method === "POST" || method === "PATCH" || method === "DELETE") {
          const apiIdx = path.indexOf("/_api/");
          const sitePath = apiIdx >= 0 ? path.substring(0, apiIdx) || "/" : "/";
          const contextUrl = location.origin + sitePath + "/_api/contextinfo";
          const cr = await fetch(contextUrl, {
            method: "POST",
            credentials: "include",
            headers: { Accept: "application/json;odata=nometadata" }
          });
          if (cr.ok) {
            const cj = await cr.json();
            digest = cj.FormDigestValue || null;
          }
        }
        const opts = {
          method,
          credentials: "include",
          headers: { Accept: "application/json;odata=nometadata" }
        };
        if (digest) opts.headers["X-RequestDigest"] = digest;
        if (body != null && method !== "GET") {
          opts.headers["Content-Type"] = "application/json;odata=nometadata";
          opts.body = typeof body === "string" ? body : JSON.stringify(body);
        }
        const r = await fetch(location.origin + path, opts);
        const text = await r.text();
        let data = text;
        try { data = JSON.parse(text); } catch (_) {}
        sendResponse(r.ok ? { ok: true, data } : { ok: false, status: r.status, error: data });
      } catch (e) {
        sendResponse({ ok: false, error: e && e.message ? e.message : String(e) });
      }
    })();
    return true;
  }

  if (message.action === "checkListPage") {
    injectAndWait(
      "checkListPage.js",
      "SPCSVListPageCheck",
      (data) => ({ isListPage: !!data.isListPage }),
      sendResponse,
      { timeoutMs: 3000, errorPayload: { isListPage: false } }
    );
    return true;
  }

  if (message.action === "getListType") {
    injectAndWait(
      "getListType.js",
      "SPCSVListTypeResult",
      (data) => ({
        isListPage: !!data.isListPage,
        isLibrary: !!data.isLibrary
      }),
      sendResponse,
      { timeoutMs: 6000, errorPayload: { isListPage: false, isLibrary: false } }
    );
    return true;
  }

  if (message.action === "getRefinableMappings") {
    injectAndWait(
      "getRefinableMappings.js",
      "SPCSVRefinableMappingsResult",
      (data) => ({
        ok: !!data.ok,
        error: data.error || null,
        mappings: Array.isArray(data.mappings) ? data.mappings : [],
        alias: data.alias != null ? data.alias : null
      }),
      sendResponse,
      {
        beforeInject() {
          let el = document.getElementById("sp-refinable-params");
          if (el) el.remove();
          el = document.createElement("script");
          el.id = "sp-refinable-params";
          el.type = "application/json";
          el.textContent = JSON.stringify({
            siteUrl: message.siteUrl || "",
            propertyName: message.propertyName || ""
          });
          (document.head || document.documentElement).appendChild(el);
        },
        timeoutMs: 15000,
        errorPayload: { ok: false, error: "Timeout or load failed", mappings: [], alias: null }
      }
    );
    return true;
  }

  if (message.action === "getSearchSchema") {
    injectAndWait(
      "getSearchSchema.js",
      "SPCSVSearchSchemaResult",
      (data) => data.error ? { ok: false, error: data.error } : { ok: true, columns: data.columns || [], siteUrl: data.siteUrl || "", listId: data.listId || "" },
      sendResponse,
      { errorPayload: { ok: false, error: "Timeout loading columns" } }
    );
    return true;
  }

  if (message.action === "getColumnCreatorContext") {
    injectAndWait(
      "getColumnCreatorContext.js",
      "SPCSVColumnCreatorContextResult",
      (data) =>
        data.ok === false
          ? { ok: false, error: data.error || "Failed to load column context.", siteFields: [], listFields: [], siteUrl: data.siteUrl || "", listId: data.listId || "" }
          : {
              ok: true,
              siteFields: data.siteFields || [],
              listFields: data.listFields || [],
              siteUrl: data.siteUrl || "",
              listId: data.listId || "",
            },
      sendResponse,
      { errorPayload: { ok: false, error: "Timeout loading column context", siteFields: [], listFields: [], siteUrl: "", listId: "" } }
    );
    return true;
  }

  if (message.action === "createColumn") {
    injectAndWait(
      "createColumn.js",
      "SPCSVCreateColumnResult",
      (data) => (data.ok ? { ok: true, internalName: data.internalName || "", title: data.title || "", type: data.type || "" } : { ok: false, error: data.error || "Create failed." }),
      sendResponse,
      {
        beforeInject() {
          attachSpColumnCreateParamsScript(message);
        },
        timeoutMs: 60000,
        errorPayload: { ok: false, error: "Creating the column timed out." },
      }
    );
    return true;
  }

  if (message.action === "getColumns") {
    injectAndWait(
      "getFields.js",
      "SPCSVFieldsResult",
      (data) => data.error ? { ok: false, error: data.error } : { ok: true, ...data },
      sendResponse,
      { errorPayload: { ok: false, error: "Timeout loading columns" } }
    );
    return true;
  }

  if (message.action === "getViewsData") {
    injectAndWait(
      "getViewsData.js",
      "SPCSVViewsDataResult",
      (data) => data.error ? { ok: false, error: data.error } : { ok: true, views: data.views || [], fields: data.fields || [], listId: data.listId, listTitle: data.listTitle, viewDetails: data.viewDetails },
      sendResponse,
      {
        beforeInject() {
          attachSpViewsParamsScript(message);
        },
        errorPayload: { ok: false, error: "Timeout loading views data" }
      }
    );
    return true;
  }

  if (message.action === "getReportListPlan") {
    injectAndWait(
      "reportListPlan.js",
      "SPReportListPlanResult",
      (data) => data.error ? { ok: false, error: data.error } : { ok: true, entries: data.entries || [] },
      sendResponse,
      {
        timeoutMs: 120000,
        beforeInject() {
          let el = document.getElementById("sp-report-list-plan-params-json");
          if (el) el.remove();
          el = document.createElement("script");
          el.id = "sp-report-list-plan-params-json";
          el.type = "application/json";
          el.textContent = JSON.stringify({
            siteUrl: message.siteUrl || "",
            includeSubsites: message.includeSubsites !== false,
            librariesOnly: message.librariesOnly === true
          });
          (document.head || document.documentElement).appendChild(el);
        },
        errorPayload: { ok: false, error: "Timeout loading list inventory" }
      }
    );
    return true;
  }

  if (message.action === "getMatrixScanPlan") {
    injectAndWait(
      "matrixScanPlan.js",
      "SPMatrixScanPlanResult",
      (data) => data.error ? { ok: false, error: data.error } : { ok: true, plan: data.plan || [] },
      sendResponse,
      {
        timeoutMs: 120000,
        beforeInject() {
          let el = document.getElementById("sp-matrix-scan-params-json");
          if (el) el.remove();
          el = document.createElement("script");
          el.id = "sp-matrix-scan-params-json";
          el.type = "application/json";
          el.textContent = JSON.stringify({
            siteUrl: message.siteUrl || "",
            includeSubsites: message.includeSubsites !== false
          });
          (document.head || document.documentElement).appendChild(el);
        },
        errorPayload: { ok: false, error: "Timeout loading site scan plan" }
      }
    );
    return true;
  }

  if (message.action === "runExportCSV") {
    chrome.runtime.sendMessage({ type: "SPCSVStartExportWorker", exportMessage: message }, function (resp) {
      if (chrome.runtime.lastError) {
        sendResponse({ ok: false, error: chrome.runtime.lastError.message });
        return;
      }
      if (!resp || !resp.ok) {
        sendResponse(resp || { ok: false, error: "Failed to start export" });
        return;
      }
      sendResponse(resp);
    });
    return true;
  }
}

function toolkitActionPromise(message) {
  return new Promise(function (resolve) {
    const asyncRet = dispatchToolkitMessage(message, resolve);
    if (asyncRet !== true) {
      queueMicrotask(function () {
        resolve({ ok: false, error: "not_handled" });
      });
    }
  });
}

let compassUrlWatchTimer = null;
let compassLastTrackedHref = "";

function refreshCompassForNavigation() {
  const panel = document.getElementById(LISTS_LAUNCHER_PANEL_ID);
  if (!panel) return;
  pageContextFullCacheHref = "";
  pageContextFullCacheData = null;
  if (typeof window.SPOT_clearCompassPagePropsCache === "function") {
    window.SPOT_clearCompassPagePropsCache(location.href);
  }
  try {
    delete panel.dataset.compassContextLoaded;
    panel._compassContextFlat = null;
  } catch (_) {}
  toolkitActionPromise({ action: "checkListPage" }).then(function (chk) {
    if (typeof window.SPOT_applyCompassTabAvailability === "function") {
      window.SPOT_applyCompassTabAvailability(panel, !!(chk && chk.isListPage));
    }
  }).catch(function () {});
  const repHost = panel.querySelector(".sp-toolkit-compass-reports-host");
  if (repHost && typeof repHost._spotRefreshReportsContext === "function") {
    void repHost._spotRefreshReportsContext();
  }
}

function startCompassUrlWatch() {
  compassLastTrackedHref = location.href;
  if (compassUrlWatchTimer) return;
  compassUrlWatchTimer = setInterval(function () {
    if (!listsLauncherExpanded) {
      stopCompassUrlWatch();
      return;
    }
    if (location.href === compassLastTrackedHref) return;
    compassLastTrackedHref = location.href;
    refreshCompassForNavigation();
  }, 750);
}

function stopCompassUrlWatch() {
  if (!compassUrlWatchTimer) return;
  clearInterval(compassUrlWatchTimer);
  compassUrlWatchTimer = null;
}

chrome.runtime.onMessage.addListener(function (message, _sender, sendResponse) {
  return dispatchToolkitMessage(message, sendResponse);
});

if (document.body) {
  initListsLauncher();
  tryResumeMatrixExportInTab();
} else {
  document.addEventListener("DOMContentLoaded", function () {
    initListsLauncher();
    tryResumeMatrixExportInTab();
  });
}

try {
  chrome.runtime.sendMessage({ type: "SPCSVExportProgressGet" }, function (data) {
    if (chrome.runtime.lastError || !data || !data.active) return;
    if (data.report === "permissionsMatrix" && data.exportParams && data.workerTabId != null) {
      chrome.tabs.getCurrent(function (tab) {
        if (chrome.runtime.lastError || !tab || tab.id !== data.workerTabId) return;
        tryResumeMatrixExportInTab(data);
      });
    }
    chrome.runtime.sendMessage({ type: "SPCSVExportEnsureWorker" });
  });
} catch (_) {}
