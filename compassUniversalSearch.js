/* global chrome */
/**
 * "Find anything" command palette for the in-page compass panel — parity with popup universal search.
 * Depends on window.SPOT_compassNavigateToPane from compassToolkitPanels (set during SPOT_initCompassToolkitPanels).
 */
(function () {
  function escapeHtml(s) {
    const d = document.createElement("div");
    d.textContent = s == null ? "" : String(s);
    return d.innerHTML;
  }

  function normalizeSearchText(v) {
    return String(v || "").toLowerCase().trim();
  }

  function normalizeGuidCopyText(value) {
    const raw = String(value == null ? "" : value).trim();
    const noBraces = raw.replace(/[{}]/g, "");
    return /^[0-9a-fA-F-]{36}$/.test(noBraces) ? noBraces : raw;
  }

  function isEmptyUniversalSearchPropValue(v) {
    if (v == null) return true;
    return String(v).trim() === "";
  }

  function parseContextFromUrl(pageUrl) {
    if (!pageUrl) return { webAbsoluteUrl: "", pageListId: "", viewId: "" };
    try {
      const u = new URL(pageUrl);
      let webAbsoluteUrl = "";
      let pageListId = "";
      let viewId = "";
      const listParam = u.searchParams.get("List") || u.searchParams.get("list");
      if (listParam) pageListId = String(listParam).replace(/^\{|\}$/g, "").replace(/%7B|%7D/gi, "");
      const viewParam = u.searchParams.get("View") || u.searchParams.get("view");
      if (viewParam) viewId = String(viewParam).replace(/^\{|\}$/g, "").replace(/%7B|%7D/gi, "");
      const path = decodeURIComponent(u.pathname.replace(/\/$/, ""));
      const segments = path.split("/").filter(Boolean);
      let siteEnd = -1;
      for (let i = 0; i < segments.length; i++) {
        if (["sites", "site", "teams"].indexOf(segments[i]) >= 0 && segments[i + 1]) {
          siteEnd = i + 1;
          break;
        }
      }
      webAbsoluteUrl = siteEnd >= 0 ? u.origin + "/" + segments.slice(0, siteEnd + 1).join("/") : u.origin;
      return { webAbsoluteUrl: webAbsoluteUrl, pageListId: pageListId, viewId: viewId };
    } catch (_) {
      return { webAbsoluteUrl: "", pageListId: "", viewId: "" };
    }
  }

  function flattenContext(obj, prefix) {
    if (obj === null || typeof obj !== "object") return prefix ? { [prefix]: obj } : {};
    const out = {};
    for (let ki = 0; ki < Object.keys(obj).length; ki++) {
      const k = Object.keys(obj)[ki];
      const v = obj[k];
      const key = prefix ? prefix + "." + k : k;
      if (v !== null && typeof v === "object" && !Array.isArray(v) && Object.prototype.toString.call(v) === "[object Object]") {
        Object.assign(out, flattenContext(v, key));
      } else {
        out[key] = v !== null && typeof v === "object" ? JSON.stringify(v) : v;
      }
    }
    return out;
  }

  function isToolkitSharePointPage(url) {
    if (!url || typeof url !== "string") return false;
    if (!url.includes("sharepoint.com")) return false;
    try {
      return !new URL(url).hostname.toLowerCase().endsWith("-admin.sharepoint.com");
    } catch (_) {
      return true;
    }
  }

  function normalizeTrailingSlash(url) {
    if (!url || typeof url !== "string") return "";
    return url.replace(/\/$/, "");
  }

  function isTenantAdminQuickLinkHref(href) {
    try {
      const u = new URL(href);
      if (u.protocol !== "https:") return false;
      return u.hostname.toLowerCase().endsWith("-admin.sharepoint.com");
    } catch (_) {
      return false;
    }
  }

  function injectCompassUniversalSearchStyles() {
    const cssText =
      "#sp-toolkit-lists-panel .compass-universal-search-overlay{position:absolute;inset:0;z-index:40;display:none;background:rgba(15,17,23,.55);backdrop-filter:blur(2px);align-items:flex-start;justify-content:center;padding:14px 12px;border-radius:inherit;overflow:hidden;box-sizing:border-box;}" +
      "#sp-toolkit-lists-panel .compass-universal-search-overlay.open{display:flex!important;}" +
      "#sp-toolkit-lists-panel .compass-universal-search-dialog{width:min(540px,calc(100% - 8px));max-height:min(560px,calc(100% - 28px));background:var(--surface);border:1.5px solid var(--border);border-radius:12px;box-shadow:0 14px 34px rgba(0,0,0,.22);overflow:hidden;display:flex;flex-direction:column;min-height:0;}" +
      "#sp-toolkit-lists-panel .compass-universal-search-dialog .compass-universal-search-input-wrap{display:flex;align-items:center;gap:12px;padding:12px 14px;border-bottom:1px solid var(--border);background:var(--surface-2);}" +
      "#sp-toolkit-lists-panel .compass-universal-search-icon{width:20px;height:20px;min-width:20px;flex-shrink:0;display:block;overflow:visible;color:var(--txt-mid);opacity:.92;pointer-events:none;}" +
      "#sp-toolkit-lists-panel .compass-universal-search-icon path{fill:currentColor;}" +
      "#sp-toolkit-lists-panel #compassUniversalSearchInput{flex:1;border:none;outline:none;background:transparent;color:var(--txt);font-family:inherit;font-size:.85rem;}" +
      "#sp-toolkit-lists-panel .compass-universal-search-hint{display:inline-flex;align-items:center;justify-content:center;min-width:28px;height:22px;padding:0 8px;font-size:.72rem;font-weight:600;line-height:1;color:var(--txt-mid);background:var(--bg);border:1px solid var(--border);border-radius:7px;}" +
      "#sp-toolkit-lists-panel #compassUniversalSearchResults{flex:1 1 auto;min-height:0;overflow-y:auto;padding:6px 0 10px;-webkit-mask-image:linear-gradient(to bottom,#000 0,#000 calc(100% - 2.75rem),rgba(0,0,0,.65) calc(100% - 1.25rem),transparent 100%);mask-image:linear-gradient(to bottom,#000 0,#000 calc(100% - 2.75rem),rgba(0,0,0,.65) calc(100% - 1.25rem),transparent 100%);-webkit-mask-size:100% 100%;mask-size:100% 100%;-webkit-mask-repeat:no-repeat;mask-repeat:no-repeat;}" +
      "#sp-toolkit-lists-panel #compassUniversalSearchResults .us-group{margin-top:6px;}" +
      "#sp-toolkit-lists-panel #compassUniversalSearchResults .us-group-title{font-size:.64rem;font-weight:700;color:var(--txt-soft);text-transform:uppercase;letter-spacing:.06em;padding:6px 12px 4px;}" +
      "#sp-toolkit-lists-panel #compassUniversalSearchResults .us-item{width:100%;text-align:left;border:none;background:transparent;color:var(--txt);cursor:pointer;padding:8px 12px;display:flex;align-items:baseline;gap:8px;}" +
      "#sp-toolkit-lists-panel #compassUniversalSearchResults .us-item:hover,#sp-toolkit-lists-panel #compassUniversalSearchResults .us-item.active{background:var(--brand-faint);}" +
      "#sp-toolkit-lists-panel #compassUniversalSearchResults .us-item-title{font-size:.8rem;font-weight:600;}" +
      "#sp-toolkit-lists-panel #compassUniversalSearchResults .us-item-meta{font-size:.68rem;color:var(--txt-soft);}" +
      "#sp-toolkit-lists-panel #compassUniversalSearchResults .us-empty{padding:16px 12px 10px;font-size:.78rem;color:var(--txt-soft);}";
    let style = document.getElementById("sp-compass-universal-search-style");
    if (!style) {
      style = document.createElement("style");
      style.id = "sp-compass-universal-search-style";
      document.head.appendChild(style);
    }
    style.textContent = cssText;
  }

  function migrateOverlaySearchGlyph(overlay) {
    if (!overlay) return;
    const wrap = overlay.querySelector(".compass-universal-search-input-wrap");
    const svg = wrap && wrap.querySelector("svg");
    if (!svg || svg.classList.contains("compass-universal-search-icon")) return;
    svg.className = "compass-universal-search-icon";
    svg.setAttribute("viewBox", "0 0 24 24");
    svg.setAttribute("aria-hidden", "true");
    svg.setAttribute("focusable", "false");
    svg.removeAttribute("stroke");
    svg.removeAttribute("fill");
    svg.innerHTML =
      "<path fill-rule=\"evenodd\" clip-rule=\"evenodd\" d=\"M10.5 3.75a6.75 6.75 0 100 13.5 6.75 6.75 0 000-13.5zM2.25 10.5a8.25 8.25 0 1114.59 5.31l3.94 3.94a.75.75 0 11-1.06 1.06l-3.94-3.94A8.25 8.25 0 012.25 10.5z\"/>";
  }

  function ensureOverlay(panel) {
    injectCompassUniversalSearchStyles();
    let overlay = panel.querySelector("#compassUniversalSearchOverlay");
    if (overlay) {
      migrateOverlaySearchGlyph(overlay);
      return overlay;
    }
    overlay = document.createElement("div");
    overlay.id = "compassUniversalSearchOverlay";
    overlay.className = "compass-universal-search-overlay";
    overlay.setAttribute("aria-hidden", "true");
    overlay.innerHTML =
      "<div class=\"compass-universal-search-dialog\" role=\"dialog\" aria-modal=\"true\" aria-label=\"Find anything\">" +
      "<div class=\"compass-universal-search-input-wrap\">" +
      "<svg class=\"compass-universal-search-icon\" viewBox=\"0 0 24 24\" aria-hidden=\"true\" focusable=\"false\">" +
      "<path fill-rule=\"evenodd\" clip-rule=\"evenodd\" d=\"M10.5 3.75a6.75 6.75 0 100 13.5 6.75 6.75 0 000-13.5zM2.25 10.5a8.25 8.25 0 1114.59 5.31l3.94 3.94a.75.75 0 11-1.06 1.06l-3.94-3.94A8.25 8.25 0 012.25 10.5z\"/>" +
      "</svg>" +
      "<input id=\"compassUniversalSearchInput\" type=\"text\" placeholder=\"Search settings and toolkit tools...\" autocomplete=\"off\" />" +
      "<span class=\"compass-universal-search-hint\">Esc</span>" +
      "</div>" +
      "<div id=\"compassUniversalSearchResults\" class=\"universal-search-results\"></div>" +
      "</div>";
    panel.appendChild(overlay);
    return overlay;
  }

  let compassUniversalSearchIndex = [];
  let compassUniversalSearchFiltered = [];
  let compassUniversalSearchActive = -1;

  function openExternalUrl(url) {
    if (!url) return;
    try {
      if (typeof chrome !== "undefined" && chrome.tabs && chrome.tabs.create) {
        chrome.tabs.create({ url: url });
        return;
      }
    } catch (_) {}
    window.open(url, "_blank", "noopener,noreferrer");
  }

  async function refreshCompassUniversalSearchIndex(panel, invoke, ctx) {
    compassUniversalSearchIndex = [];
    const siteUrl = (ctx && ctx.siteUrl) || "";
    const pageHref = typeof location !== "undefined" ? location.href : "";
    const onSp = pageHref && isToolkitSharePointPage(pageHref);

    let pageCtx = null;
    let parsedCtx = { webAbsoluteUrl: "", pageListId: "", viewId: "" };
    let isListPageForSearch = false;
    let schemaColumns = [];

    if (onSp) {
      parsedCtx = parseContextFromUrl(pageHref);
      try {
        pageCtx = await invoke({ action: "getPageContext" });
      } catch (_) {}
      try {
        const chkRes = await invoke({ action: "checkListPage" });
        isListPageForSearch = !!(chkRes && chkRes.isListPage);
      } catch (_) {}
      try {
        const schemaRes = await invoke({ action: "getSearchSchema" });
        if (schemaRes && schemaRes.ok && Array.isArray(schemaRes.columns)) schemaColumns = schemaRes.columns;
      } catch (_) {}
    }

    const prefetch = window.SPOT_prefetchCompassUniversalSearchData;
    if (typeof prefetch === "function") {
      try {
        await prefetch(panel, invoke, ctx || {});
      } catch (_) {}
    }

    let flatProps = null;
    try {
      const full = await invoke({ action: "getPageContextFull" });
      if (full && full.ok && full.data && typeof full.data === "object") flatProps = flattenContext(full.data);
    } catch (_) {}
    try {
      const pj = await invoke({ action: "getPageContextJson" });
      if (pj && pj.ok && pj.data && typeof pj.data === "object") {
        const jf = flattenContext(pj.data);
        flatProps = flatProps && Object.keys(flatProps).length ? Object.assign({}, jf, flatProps) : jf;
      }
    } catch (_) {}
    if (pageCtx && pageCtx.ok && pageCtx.webAbsoluteUrl) {
      const baseFlat = {
        webAbsoluteUrl: pageCtx.webAbsoluteUrl,
        siteAbsoluteUrl: pageCtx.siteAbsoluteUrl || "",
        pageListId: pageCtx.pageListId || "",
        listUrl: pageCtx.listUrl || "",
        siteId: pageCtx.siteId || "",
        webId: pageCtx.webId || ""
      };
      flatProps = flatProps && Object.keys(flatProps).length ? Object.assign({}, flatProps, baseFlat) : baseFlat;
    }

    function addItem(group, title, meta, keywords, run) {
      compassUniversalSearchIndex.push({
        group: group,
        title: title,
        meta: meta,
        keywords: normalizeSearchText(keywords),
        run: run
      });
    }

    panel.querySelectorAll(".sp-toolkit-compass-pane-tab").forEach(function (tab) {
      if (tab.disabled) return;
      const id = tab.getAttribute("data-compass-pane") || "";
      const title = (tab.textContent || "").trim();
      if (!id || !title) return;
      addItem("Tabs", title, "Switch tab", title + " " + id, function () {
        const nav = window.SPOT_compassNavigateToPane;
        if (typeof nav === "function") nav(id);
      });
    });

    if (onSp) {
      const siteBase = (parsedCtx.webAbsoluteUrl || "").replace(/\/$/, "") || new URL(pageHref).origin;
      const jumpToContextValue = function (valueText, fallbackText) {
        const nav = window.SPOT_compassNavigateToPane;
        const finish = function () {
          const input = panel.querySelector("#contextFilterInput");
          const byValue = panel.querySelector("#contextFilterByValue");
          if (byValue) byValue.checked = true;
          if (input) {
            input.value = String(valueText || fallbackText || "");
            input.dispatchEvent(new Event("input", { bubbles: true }));
          }
          try {
            if (valueText) navigator.clipboard.writeText(normalizeGuidCopyText(String(valueText)));
          } catch (_) {}
        };
        if (panel.dataset.compassContextLoaded === "1") {
          if (typeof nav === "function") nav("context").then(finish).catch(finish);
        } else if (typeof nav === "function") {
          nav("context").then(finish).catch(finish);
        } else finish();
      };

      const urlPath = (function () {
        try {
          return new URL(pageHref).pathname || "/";
        } catch (_) {
          return "";
        }
      })();
      const listId =
        (pageCtx && (pageCtx.pageListId || pageCtx.listId || pageCtx.listGuid)) || parsedCtx.pageListId || "";
      const viewId = (pageCtx && (pageCtx.viewId || pageCtx.pageViewId)) || parsedCtx.viewId || "";
      const webUrl = (pageCtx && pageCtx.webAbsoluteUrl) || parsedCtx.webAbsoluteUrl || "";
      const siteUrlVal = (pageCtx && pageCtx.siteAbsoluteUrl) || "";
      const siteIdVal = (pageCtx && pageCtx.siteId) || "";

      addItem("Current Page", "Page URL", urlPath, "page url current url location " + pageHref + " " + urlPath, function () {
        jumpToContextValue(pageHref, "url");
      });
      if (webUrl) {
        addItem("Current Page", "Web URL", webUrl, "web url webabsoluteurl " + webUrl, function () {
          jumpToContextValue(webUrl, "webAbsoluteUrl");
        });
      }
      if (siteUrlVal) {
        addItem("Current Page", "Site URL", siteUrlVal, "site url siteabsoluteurl " + siteUrlVal, function () {
          jumpToContextValue(siteUrlVal, "siteAbsoluteUrl");
        });
      }
      if (listId) {
        addItem("Current Page", "List ID", String(listId), "list id listid pagelistid listguid guid " + listId, function () {
          jumpToContextValue(listId, "listid");
        });
      }
      if (viewId) {
        addItem("Current Page", "View ID", String(viewId), "view id viewid pageviewid guid " + viewId, function () {
          jumpToContextValue(viewId, "viewid");
        });
      }
      if (siteIdVal) {
        addItem("Current Page", "Site ID", String(siteIdVal), "site id siteid guid " + siteIdVal, function () {
          jumpToContextValue(siteIdVal, "siteid");
        });
      }

      const qc = panel.querySelector("#quicklinksContent");
      if (qc) {
        qc.querySelectorAll(".sp-toolkit-ql-section").forEach(function (section) {
          const sectionLabel = (section.querySelector(".sp-toolkit-ql-section-label") || {}).textContent || "";
          section.querySelectorAll("a[href]").forEach(function (a) {
            const label = (a.textContent || "").trim();
            const href = a.getAttribute("href") || "";
            if (!label || !href) return;
            addItem(
              "Quick Links",
              label,
              sectionLabel.trim() || "Quick link",
              label + " " + href + " " + sectionLabel,
              function () {
                if (/second stage recycle bin/i.test(label) && /adminrecyclebin\.aspx/i.test(href) && siteBase) {
                  const firstUrl = normalizeTrailingSlash(siteBase) + "/_layouts/15/RecycleBin.aspx";
                  try {
                    chrome.runtime.sendMessage(
                      {
                        type: "SPOToolkitSecondStageRecycleBin",
                        firstUrl: firstUrl,
                        secondUrl: href
                      },
                      function () {
                        void chrome.runtime.lastError;
                      }
                    );
                  } catch (_) {}
                } else if (isTenantAdminQuickLinkHref(href)) {
                  try {
                    chrome.runtime.sendMessage({ type: "SPOToolkitOpenTenantAdminWithWait", url: href }, function () {
                      void chrome.runtime.lastError;
                    });
                  } catch (_) {}
                } else {
                  openExternalUrl(href);
                }
              }
            );
          });
        });
      }

      if (flatProps && typeof flatProps === "object") {
        Object.entries(flatProps).slice(0, 200).forEach(function (kv) {
          const k = kv[0];
          const v = kv[1];
          if (isEmptyUniversalSearchPropValue(v)) return;
          const val = String(v);
          addItem("Page Props", k, val.slice(0, 90), k + " " + val, function () {
            const nav = window.SPOT_compassNavigateToPane;
            const apply = function () {
              const input = panel.querySelector("#contextFilterInput");
              const byValue = panel.querySelector("#contextFilterByValue");
              if (byValue) byValue.checked = false;
              if (input) {
                input.value = k;
                input.dispatchEvent(new Event("input", { bubbles: true }));
              }
            };
            if (typeof nav === "function") nav("context").then(apply).catch(apply);
            else apply();
          });
        });
      }

      if (schemaColumns.length) {
        schemaColumns.slice(0, 300).forEach(function (c) {
          const title = c.title || c.internalName || "";
          const internal = c.internalName || "";
          const crawled = c.crawledProperty || "";
          const type = c.type || "";
          addItem(
            "Columns",
            title,
            internal || type,
            title + " " + internal + " " + crawled + " " + type,
            function () {
              const nav = window.SPOT_compassNavigateToPane;
              const apply = function () {
                const input = panel.querySelector("#searchSchemaFilter");
                if (input) {
                  input.value = title || internal;
                  input.dispatchEvent(new Event("input", { bubbles: true }));
                }
              };
              if (typeof nav === "function") nav("columns").then(apply).catch(apply);
              else apply();
            }
          );
        });
      }

      if (isListPageForSearch) {
        addItem("Tools", "View Manager", "List & library views", "view manager views manage edit editor view editor", function () {
          try {
            chrome.runtime.sendMessage({ type: "SPOToolkitOpenViewsPage" }, function () {
              void chrome.runtime.lastError;
            });
          } catch (_) {
            openExternalUrl(chrome.runtime.getURL("views.html"));
          }
        });
        addItem("Tools", "View formatter", "JSON editor & list preview", "view formatter json column formatting format", function () {
          try {
            chrome.runtime.sendMessage({ type: "SPOToolkitOpenViewFormatter", previewUrl: pageHref }, function () {
              void chrome.runtime.lastError;
            });
          } catch (_) {
            const u = new URL(chrome.runtime.getURL("view-formatter.html"));
            u.searchParams.set("src", pageHref);
            openExternalUrl(u.toString());
          }
        });
        addItem("Reports", "Run Export", "Reports", "run export reports", function () {
          const nav = window.SPOT_compassNavigateToPane;
          const apply = function () {
            const btn = panel.querySelector(".sp-toolkit-btn-run-export");
            if (btn) btn.focus();
          };
          if (typeof nav === "function") nav("reports").then(apply).catch(apply);
          else apply();
        });
        addItem("Reports", "Permissions Matrix", "Report type", "permissions matrix", function () {
          const nav = window.SPOT_compassNavigateToPane;
          const apply = function () {
            const sel = panel.querySelector(".sp-toolkit-compass-report-select");
            if (sel) {
              sel.value = "permissionsMatrix";
              sel.dispatchEvent(new Event("change", { bubbles: true }));
            }
            const btn = panel.querySelector(".sp-toolkit-btn-run-export");
            if (btn) btn.focus();
          };
          if (typeof nav === "function") nav("reports").then(apply).catch(apply);
          else apply();
        });
        addItem("Reports", "Path Length Report", "Report type", "path length", function () {
          const nav = window.SPOT_compassNavigateToPane;
          const apply = function () {
            const sel = panel.querySelector(".sp-toolkit-compass-report-select");
            if (sel) {
              sel.value = "pathLengths";
              sel.dispatchEvent(new Event("change", { bubbles: true }));
            }
            const btn = panel.querySelector(".sp-toolkit-btn-run-export");
            if (btn) btn.focus();
          };
          if (typeof nav === "function") nav("reports").then(apply).catch(apply);
          else apply();
        });
        addItem("Reports", "Folder and Item Counts", "Report type", "folder item count", function () {
          const nav = window.SPOT_compassNavigateToPane;
          const apply = function () {
            const sel = panel.querySelector(".sp-toolkit-compass-report-select");
            if (sel) {
              sel.value = "folderCount";
              sel.dispatchEvent(new Event("change", { bubbles: true }));
            }
            const btn = panel.querySelector(".sp-toolkit-btn-run-export");
            if (btn) btn.focus();
          };
          if (typeof nav === "function") nav("reports").then(apply).catch(apply);
          else apply();
        });
      }
    }
  }

  function renderCompassUniversalSearchResults(inputEl, resultsEl, queryRaw) {
    if (!resultsEl) return;
    const query = normalizeSearchText(queryRaw);
    compassUniversalSearchFiltered = query
      ? compassUniversalSearchIndex.filter(function (it) {
          return it.keywords.indexOf(query) >= 0;
        })
      : compassUniversalSearchIndex;
    compassUniversalSearchActive = compassUniversalSearchFiltered.length ? 0 : -1;

    if (!compassUniversalSearchFiltered.length) {
      resultsEl.innerHTML = "<div class=\"us-empty\">No matches yet.</div>";
      return;
    }

    const groups = new Map();
    compassUniversalSearchFiltered.forEach(function (it) {
      if (!groups.has(it.group)) groups.set(it.group, []);
      groups.get(it.group).push(it);
    });

    resultsEl.innerHTML = "";
    let flatIndex = 0;
    groups.forEach(function (items, group) {
      const g = document.createElement("div");
      g.className = "us-group";
      const h = document.createElement("div");
      h.className = "us-group-title";
      h.textContent = group;
      g.appendChild(h);
      items.slice(0, 18).forEach(function (it) {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "us-item" + (flatIndex === compassUniversalSearchActive ? " active" : "");
        btn.dataset.idx = String(flatIndex);
        btn.innerHTML =
          "<span class=\"us-item-title\">" + escapeHtml(it.title) + "</span>" +
          (it.meta ? "<span class=\"us-item-meta\">" + escapeHtml(it.meta) + "</span>" : "");
        btn.addEventListener("click", function () {
          it.run();
          const ov = document.getElementById("compassUniversalSearchOverlay");
          closeCompassUniversalSearch(ov);
        });
        g.appendChild(btn);
        flatIndex += 1;
      });
      resultsEl.appendChild(g);
    });
  }

  function moveCompassUniversalSearchSelection(resultsEl, dir) {
    if (!compassUniversalSearchFiltered.length) return;
    const max = compassUniversalSearchFiltered.length - 1;
    compassUniversalSearchActive = Math.max(0, Math.min(max, compassUniversalSearchActive + dir));
    const buttons = resultsEl ? Array.from(resultsEl.querySelectorAll(".us-item")) : [];
    buttons.forEach(function (b, idx) {
      b.classList.toggle("active", idx === compassUniversalSearchActive);
    });
    const activeEl = buttons[compassUniversalSearchActive];
    if (activeEl) activeEl.scrollIntoView({ block: "nearest" });
  }

  function closeCompassUniversalSearch(overlay) {
    if (!overlay) return;
    overlay.classList.remove("open");
    overlay.setAttribute("aria-hidden", "true");
    const input = overlay.querySelector("#compassUniversalSearchInput");
    if (input) input.value = "";
  }

  function buildTabsOnlyCompassSearchIndex(panel) {
    compassUniversalSearchIndex = [];
    function addItem(group, title, meta, keywords, run) {
      compassUniversalSearchIndex.push({
        group: group,
        title: title,
        meta: meta,
        keywords: normalizeSearchText(keywords),
        run: run
      });
    }
    panel.querySelectorAll(".sp-toolkit-compass-pane-tab").forEach(function (tab) {
      if (tab.disabled) return;
      const id = tab.getAttribute("data-compass-pane") || "";
      const title = (tab.textContent || "").trim();
      if (!id || !title) return;
      addItem("Tabs", title, "Switch tab", title + " " + id, function () {
        const nav = window.SPOT_compassNavigateToPane;
        if (typeof nav === "function") nav(id);
      });
    });
  }

  function focusCompassUniversalSearchInput(input) {
    if (!input) return;
    function go() {
      try {
        input.focus({ preventScroll: true });
      } catch (_) {
        input.focus();
      }
      try {
        input.setSelectionRange(0, 0);
      } catch (_) {}
    }
    go();
    setTimeout(go, 0);
    requestAnimationFrame(function () {
      requestAnimationFrame(go);
    });
  }

  function openCompassUniversalSearch(panel, invoke, ctx) {
    injectCompassUniversalSearchStyles();
    const overlay = ensureOverlay(panel);
    migrateOverlaySearchGlyph(overlay);
    const input = overlay.querySelector("#compassUniversalSearchInput");
    const results = overlay.querySelector("#compassUniversalSearchResults");
    if (!input || !results) return;
    buildTabsOnlyCompassSearchIndex(panel);
    input.value = "";
    renderCompassUniversalSearchResults(input, results, "");
    overlay.classList.add("open");
    overlay.setAttribute("aria-hidden", "false");
    focusCompassUniversalSearchInput(input);
    refreshCompassUniversalSearchIndex(panel, invoke, ctx).then(function () {
      renderCompassUniversalSearchResults(input, results, input.value || "");
      focusCompassUniversalSearchInput(input);
    });
  }

  function wireCompassUniversalSearch(panel, invoke, ctx) {
    injectCompassUniversalSearchStyles();
    const btn = panel.querySelector("#btnCompassUniversalSearch");
    if (btn && !btn.dataset.spotCompassUsBound) {
      btn.dataset.spotCompassUsBound = "1";
      btn.addEventListener("click", function (e) {
        e.preventDefault();
        e.stopPropagation();
        openCompassUniversalSearch(panel, invoke, ctx);
      });
    }

    const overlay = panel.querySelector("#compassUniversalSearchOverlay") || ensureOverlay(panel);
    const input = overlay.querySelector("#compassUniversalSearchInput");
    const results = overlay.querySelector("#compassUniversalSearchResults");
    if (overlay && !overlay.dataset.spotCompassUsBound) {
      overlay.dataset.spotCompassUsBound = "1";
      overlay.addEventListener("click", function (e) {
        if (e.target === overlay) closeCompassUniversalSearch(overlay);
      });
    }
    if (input && !input.dataset.spotCompassUsBound) {
      input.dataset.spotCompassUsBound = "1";
      input.addEventListener("input", function () {
        renderCompassUniversalSearchResults(input, results, input.value);
      });
      input.addEventListener("keydown", function (e) {
        if (e.key === "Escape") {
          e.preventDefault();
          closeCompassUniversalSearch(overlay);
          return;
        }
        if (e.key === "ArrowDown") {
          e.preventDefault();
          moveCompassUniversalSearchSelection(results, 1);
          return;
        }
        if (e.key === "ArrowUp") {
          e.preventDefault();
          moveCompassUniversalSearchSelection(results, -1);
          return;
        }
        if (e.key === "Enter" && compassUniversalSearchFiltered[compassUniversalSearchActive]) {
          e.preventDefault();
          compassUniversalSearchFiltered[compassUniversalSearchActive].run();
          closeCompassUniversalSearch(overlay);
        }
      });
    }

    if (!window._SPOT_compassUniversalSearchDocKeyBound) {
      window._SPOT_compassUniversalSearchDocKeyBound = true;
      document.addEventListener(
        "keydown",
        function (e) {
          const isCmdK = (e.ctrlKey || e.metaKey) && String(e.key).toLowerCase() === "k";
          const panelEl = document.getElementById("sp-toolkit-lists-panel");
          if (!panelEl || panelEl.style.display === "none") return;
          if (!panelEl.classList.contains("sp-toolkit-lists-panel-open")) return;
          const invokeFn = window.SPOT_compassUniversalSearchInvoke;
          const ctxObj = window.SPOT_compassUniversalSearchCtx || {};
          if (isCmdK) {
            e.preventDefault();
            if (typeof invokeFn !== "function") return;
            openCompassUniversalSearch(panelEl, invokeFn, ctxObj);
            return;
          }
          const ov = panelEl.querySelector("#compassUniversalSearchOverlay");
          if (e.key === "Escape" && ov && ov.classList.contains("open")) {
            e.preventDefault();
            closeCompassUniversalSearch(ov);
          }
        },
        true
      );
    }
  }

  window.SPOT_initCompassUniversalSearch = function (panel, invokeToolkitAction, ctx) {
    if (!panel || typeof invokeToolkitAction !== "function") return;
    window.SPOT_compassUniversalSearchInvoke = invokeToolkitAction;
    window.SPOT_compassUniversalSearchCtx = ctx || {};
    window.SPOT_compassOpenUniversalSearch = function () {
      openCompassUniversalSearch(panel, invokeToolkitAction, ctx || {});
    };
    wireCompassUniversalSearch(panel, invokeToolkitAction, ctx || {});
  };
})();
