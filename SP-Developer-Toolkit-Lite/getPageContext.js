(function () {
  var i = window._spPageContextInfo || {};
  window.postMessage({
    __spcsv: true,
    type: "SPCSVPageContext",
    webAbsoluteUrl: i.webAbsoluteUrl,
    siteAbsoluteUrl: i.siteAbsoluteUrl,
    pageListId: i.pageListId,
    listUrl: i.listUrl || i.listServerRelativeUrl
  }, "*");
})();
