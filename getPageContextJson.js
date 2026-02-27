// Injected into page context. Fetches current page URL with ?as=json (modern SharePoint context fallback).
(function () {
  var url = window.location.href;
  if (url.indexOf("?") >= 0) url += "&as=json";
  else url += "?as=json";
  fetch(url, { credentials: "include", headers: { Accept: "application/json" } })
    .then(function (r) { return r.json(); })
    .then(function (data) {
      window.postMessage({ __spcsv: true, type: "SPCSVPageContextJson", data: data }, "*");
    })
    .catch(function (err) {
      window.postMessage({ __spcsv: true, type: "SPCSVPageContextJson", error: (err && err.message) ? err.message : String(err) }, "*");
    });
})();
