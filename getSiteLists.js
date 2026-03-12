(function () {
  var spi = window._spPageContextInfo || {};
  var siteUrl = (spi.webAbsoluteUrl || spi.siteAbsoluteUrl || "").replace(/\/$/, "");
  var rawListId = (spi.pageListId || spi.listId || "").replace(/^\{|\}$/g, "").replace(/%7B|%7D/gi, "").trim();
  var currentListId = rawListId ? ("{" + rawListId.toUpperCase() + "}") : "";
  if (!siteUrl && typeof location !== "undefined" && location.origin && location.pathname) {
    try {
      var path = decodeURIComponent(location.pathname.replace(/\/$/, "") || "");
      var segments = path.split("/").filter(Boolean);
      var siteEnd = -1;
      for (var i = 0; i < segments.length; i++) {
        if (["sites", "site", "teams"].indexOf(segments[i]) >= 0 && segments[i + 1]) {
          siteEnd = i + 1;
          break;
        }
      }
      siteUrl = siteEnd >= 0 ? location.origin + "/" + segments.slice(0, siteEnd + 1).join("/") : location.origin;
    } catch (_) {}
  }
  if (!siteUrl) {
    window.postMessage({ __spcsv: true, type: "SPCSVSiteListsResult", error: "No site URL" }, "*");
    return;
  }
  var typeLabels = { 100: "List", 101: "Document library", 106: "Events list", 107: "Links list", 109: "Picture library", 110: "Survey", 119: "Page library", 171: "Issue tracking" };
  var accept = "application/json;odata=nometadata";
  var listsUrl = siteUrl + "/_api/web/lists?$filter=Hidden eq false&$select=Id,Title,DefaultViewUrl,BaseTemplate,ItemCount,LastItemModifiedDate&$orderby=Title asc&$top=500";
  var subsitesUrl = siteUrl + "/_api/web/webs?$select=Title,Url&$top=200";
  var webUrl = siteUrl + "/_api/web?$select=Title";
  function tenantFromHost() {
    try {
      var host = (typeof location !== "undefined" && location.hostname) ? location.hostname : "";
      var part = host.split(".")[0] || "";
      return part ? part.charAt(0).toUpperCase() + part.slice(1).toLowerCase() : "";
    } catch (_) { return ""; }
  }
  var tenantName = tenantFromHost();
  fetch(webUrl, { credentials: "include", headers: { Accept: accept } })
    .then(function (r) { return r.json(); })
    .then(function (webJson) {
      var webTitle = (webJson.Title != null ? webJson.Title : (webJson.d && webJson.d.Title)) || "";
      return fetch(listsUrl, { credentials: "include", headers: { Accept: accept } })
        .then(function (r) { return r.json(); })
        .then(function (j) {
          var results = (j.value != null ? j.value : (j.d && j.d.results) != null ? j.d.results : (j.d && j.d) != null ? j.d : []);
          if (!Array.isArray(results)) results = [];
          var base = location.origin;
          var lists = results.map(function (item) {
            var title = item.Title || item.title || "";
            var viewUrl = item.DefaultViewUrl || item.defaultViewUrl || "";
            if (viewUrl && viewUrl.indexOf("http") !== 0) viewUrl = base + (viewUrl.indexOf("/") === 0 ? viewUrl : "/" + viewUrl);
            var baseTemplate = item.BaseTemplate != null ? item.BaseTemplate : item.baseTemplate;
            var isLibrary = (baseTemplate === 101 || baseTemplate === 109 || baseTemplate === 119);
            var typeLabel = typeLabels[baseTemplate] || "List";
            var itemCount = item.ItemCount != null ? item.ItemCount : (item.itemCount != null ? item.itemCount : 0);
            var modified = item.LastItemModifiedDate || item.lastItemModifiedDate || "";
            return { id: item.Id, title: title, viewUrl: viewUrl || "", isLibrary: isLibrary, typeLabel: typeLabel, itemCount: itemCount, modified: modified };
          }).filter(function (item) { return item.viewUrl; });
          window.postMessage({ __spcsv: true, type: "SPCSVSiteListsResult", lists: lists, siteUrl: siteUrl, currentListId: currentListId || null, tenantName: tenantName, siteName: webTitle }, "*");
        });
    })
    .catch(function () {
      return fetch(listsUrl, { credentials: "include", headers: { Accept: accept } })
        .then(function (r) { return r.json(); })
        .then(function (j) {
          var results = (j.value != null ? j.value : (j.d && j.d.results) != null ? j.d.results : (j.d && j.d) != null ? j.d : []);
          if (!Array.isArray(results)) results = [];
          var base = location.origin;
          var lists = results.map(function (item) {
            var title = item.Title || item.title || "";
            var viewUrl = item.DefaultViewUrl || item.defaultViewUrl || "";
            if (viewUrl && viewUrl.indexOf("http") !== 0) viewUrl = base + (viewUrl.indexOf("/") === 0 ? viewUrl : "/" + viewUrl);
            var baseTemplate = item.BaseTemplate != null ? item.BaseTemplate : item.baseTemplate;
            var isLibrary = (baseTemplate === 101 || baseTemplate === 109 || baseTemplate === 119);
            var typeLabel = typeLabels[baseTemplate] || "List";
            var itemCount = item.ItemCount != null ? item.ItemCount : (item.itemCount != null ? item.itemCount : 0);
            var modified = item.LastItemModifiedDate || item.lastItemModifiedDate || "";
            return { id: item.Id, title: title, viewUrl: viewUrl || "", isLibrary: isLibrary, typeLabel: typeLabel, itemCount: itemCount, modified: modified };
          }).filter(function (item) { return item.viewUrl; });
          window.postMessage({ __spcsv: true, type: "SPCSVSiteListsResult", lists: lists, siteUrl: siteUrl, currentListId: currentListId || null, tenantName: tenantName, siteName: "" }, "*");
        });
    })
    .catch(function (err) {
      window.postMessage({ __spcsv: true, type: "SPCSVSiteListsResult", error: (err && err.message) || "Failed to load lists", tenantName: tenantName, siteName: "" }, "*");
    });
  fetch(subsitesUrl, { credentials: "include", headers: { Accept: accept } })
    .then(function (r) { return r.json(); })
    .then(function (j) {
      var results = (j.value != null ? j.value : (j.d && j.d.results) != null ? j.d.results : (j.d && j.d) != null ? j.d : []);
      if (!Array.isArray(results)) results = [];
      var base = location.origin;
      var subsites = results.map(function (item) {
        var title = item.Title || item.title || "";
        var url = (item.Url || item.url || "").replace(/\/$/, "");
        if (url && url.indexOf("http") !== 0) url = base + (url.indexOf("/") === 0 ? url : "/" + url);
        return { title: title, url: url };
      }).filter(function (item) { return item.url; });
      window.postMessage({ __spcsv: true, type: "SPCSVSubsitesResult", subsites: subsites }, "*");
    })
    .catch(function () {
      window.postMessage({ __spcsv: true, type: "SPCSVSubsitesResult", subsites: [] }, "*");
    });
})();
