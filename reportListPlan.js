// Injected page script: inventory lists/libraries for site + optional subsites.
(function () {
  function readParams() {
    try {
      var el = document.getElementById("sp-report-list-plan-params-json");
      if (el && el.textContent) return JSON.parse(el.textContent);
    } catch (_) {}
    return {};
  }

  var params = readParams();
  var includeSubsites = params.includeSubsites !== false;
  var librariesOnly = params.librariesOnly === true;

  function post(type, payload) {
    window.postMessage(Object.assign({ __spcsv: true, type: type }, payload || {}), "*");
  }

  function normalizePath(urlOrPath) {
    if (!urlOrPath) return "";
    var s = String(urlOrPath).trim();
    if (s.charAt(0) === "/") return s.replace(/\/$/, "") || "/";
    try { return new URL(s).pathname.replace(/\/$/, "") || "/"; } catch (_) { return s.replace(/\/$/, ""); }
  }

  function normalizeGuid(g) {
    if (g == null) return "";
    var s = String(g).trim().replace(/[{}]/g, "").toUpperCase();
    return s || "";
  }

  var apiOrigin = String(location.origin || "").replace(/\/$/, "");
  var spi = typeof window._spPageContextInfo === "object" ? window._spPageContextInfo : null;
  if (spi && spi.webAbsoluteUrl) {
    try { apiOrigin = new URL(spi.webAbsoluteUrl).origin.replace(/\/$/, ""); } catch (_) {}
  } else if (params.siteUrl) {
    try { apiOrigin = new URL(String(params.siteUrl)).origin.replace(/\/$/, ""); } catch (_) {}
  }

  function webUrlFromPath(path) {
    var p = normalizePath(path);
    if (!p || p === "/") return apiOrigin;
    return apiOrigin + p;
  }

  async function fetchJson(url, accept) {
    var r = await fetch(url, { credentials: "include", headers: { Accept: accept || "application/json;odata=nometadata" } });
    if (!r.ok) throw new Error("HTTP " + r.status + " for " + url);
    return r.json();
  }

  async function buildPlan() {
    var accept = "application/json;odata=nometadata";
    var entries = [];
    var seenWeb = {};

    async function addWebLists(path) {
      var p = normalizePath(path);
      if (!p) return;
      var webKey = p.toLowerCase();
      if (seenWeb[webKey]) return;
      seenWeb[webKey] = true;
      var webUrl = webUrlFromPath(p);
      var webJson = await fetchJson(webUrl + "/_api/web?$select=Title,ServerRelativeUrl,Url", accept);
      var siteTitle = (webJson.Title || webJson.title || p || "Site").trim();
      var sitePath = normalizePath(webJson.ServerRelativeUrl || webJson.serverRelativeUrl || p);
      var listsJson = await fetchJson(
        webUrl + "/_api/web/lists?$select=Id,Title,BaseTemplate,ItemCount,DefaultViewUrl&$filter=Hidden eq false&$orderby=Title asc&$top=5000",
        accept
      );
      var lists = listsJson.value || listsJson.results || [];
      for (var i = 0; i < lists.length; i++) {
        var lst = lists[i];
        var title = (lst.Title || lst.title || "").trim();
        if (!title || title === "Access Requests") continue;
        var baseTemplate = lst.BaseTemplate != null ? lst.BaseTemplate : lst.baseTemplate;
        var isLibrary = baseTemplate === 101 || baseTemplate === 109 || baseTemplate === 119;
        if (librariesOnly && baseTemplate !== 101) continue;
        var listId = normalizeGuid(lst.Id || lst.id);
        if (!listId) continue;
        var viewUrl = lst.DefaultViewUrl || lst.defaultViewUrl || "";
        if (viewUrl && viewUrl.indexOf("http") !== 0) {
          viewUrl = apiOrigin + (viewUrl.indexOf("/") === 0 ? viewUrl : "/" + viewUrl);
        }
        entries.push({
          siteUrl: webUrl.replace(/\/$/, ""),
          siteTitle: siteTitle,
          sitePath: sitePath,
          listId: listId,
          listTitle: title,
          baseTemplate: baseTemplate,
          isLibrary: isLibrary,
          itemCount: parseInt(lst.ItemCount || lst.itemCount || 0, 10) || 0,
          viewUrl: viewUrl || ""
        });
      }
      if (!includeSubsites) return;
      var subs = await fetchJson(webUrl + "/_api/web/webs?$select=Title,ServerRelativeUrl,Url&$top=500", accept);
      var subItems = subs.value || subs.results || [];
      for (var s = 0; s < subItems.length; s++) {
        var subPath = subItems[s].ServerRelativeUrl || subItems[s].serverRelativeUrl;
        if (!subPath && (subItems[s].Url || subItems[s].url)) subPath = normalizePath(subItems[s].Url || subItems[s].url);
        if (subPath) await addWebLists(subPath);
      }
    }

    var startPath = normalizePath(params.siteUrl || (spi && spi.webServerRelativeUrl) || "/");
    if (startPath && startPath !== "/" && startPath.indexOf("/") === 0) {
      await addWebLists(startPath);
    } else {
      var root = await fetchJson(apiOrigin + "/_api/web?$select=ServerRelativeUrl", accept);
      await addWebLists(root.ServerRelativeUrl || root.serverRelativeUrl || "/");
    }
    return entries;
  }

  buildPlan().then(function (entries) {
    post("SPReportListPlanResult", { ok: true, entries: entries });
  }).catch(function (e) {
    post("SPReportListPlanResult", { ok: false, error: (e && e.message) || String(e) });
  });
})();
