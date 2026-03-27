(function () {
  var i = window._spPageContextInfo || {};
  window.postMessage({
    __spcsv: true,
    type: "SPCSVPageContext",
    webAbsoluteUrl: i.webAbsoluteUrl,
    siteAbsoluteUrl: i.siteAbsoluteUrl,
    siteId: i.siteId || i.siteID || i.SiteId || i.SiteID,
    webId: i.webId || i.webID || i.WebId || i.WebID,
    pageListId: i.pageListId,
    listUrl: i.listUrl || i.listServerRelativeUrl
  }, "*");
})();
