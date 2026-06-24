// Injected: returns whether the page is a list/library and if it's a library (BaseTemplate 101).
(function () {
  var spi = window._spPageContextInfo || {};
  var listUrl = spi.listUrl || spi.listServerRelativeUrl;
  if (!listUrl) {
    window.postMessage({ __spcsv: true, type: "SPCSVListTypeResult", isListPage: false, isLibrary: false }, "*");
    return;
  }
  var siteUrl = (spi.webAbsoluteUrl || spi.siteAbsoluteUrl || "").replace(/\/$/, "");
  var enc = encodeURIComponent("'" + listUrl.replace(/'/g, "''") + "'");
  // Step 1: get list Id (same as getFields.js)
  fetch(siteUrl + "/_api/web/GetList(@u)/Id?@u=" + enc, {
    credentials: "include",
    headers: { "Accept": "application/json;odata=nometadata" }
  })
    .then(function (r) { return r.json(); })
    .then(function (j) {
      var listId = (j.value != null ? j.value : (j.d && j.d.Id) != null ? j.d.Id : j.Id || "");
      listId = ("" + listId).replace(/[{}]/g, "").trim();
      if (!listId) {
        window.postMessage({ __spcsv: true, type: "SPCSVListTypeResult", isListPage: true, isLibrary: false }, "*");
        return;
      }
      return fetch(siteUrl + "/_api/web/lists(guid'" + listId + "')?$select=BaseTemplate", {
        credentials: "include",
        headers: { "Accept": "application/json;odata=nometadata" }
      }).then(function (r2) { return r2.json(); }).then(function (listJ) {
        var bt = listJ.BaseTemplate != null ? listJ.BaseTemplate : (listJ.d && listJ.d.BaseTemplate != null ? listJ.d.BaseTemplate : null);
        var n = parseInt(bt, 10);
        var isLibrary = n === 101;
        window.postMessage({ __spcsv: true, type: "SPCSVListTypeResult", isListPage: true, isLibrary: isLibrary }, "*");
      });
    })
    .catch(function () {
      window.postMessage({ __spcsv: true, type: "SPCSVListTypeResult", isListPage: true, isLibrary: false }, "*");
    });
})();
