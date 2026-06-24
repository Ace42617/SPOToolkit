// Injected into page context. Posts back whether the page is a list or library (has listUrl).
(function () {
  var spi = window._spPageContextInfo || {};
  var isListPage = !!(spi.listUrl || spi.listServerRelativeUrl);
  window.postMessage({ __spcsv: true, type: "SPCSVListPageCheck", isListPage: isListPage }, "*");
})();
