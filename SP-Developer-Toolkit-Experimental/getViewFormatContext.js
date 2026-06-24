// Injected into SharePoint page: list id, optional view id, web URL for View formatter save.
(function () {
  function normGuid(s) {
    return String(s || "")
      .replace(/[{}]/g, "")
      .replace(/%7B|%7D/gi, "")
      .trim();
  }
  var spi = window._spPageContextInfo || {};
  var listId = normGuid(spi.pageListId);
  var viewId = normGuid(spi.viewId || spi.pageViewId);
  var webAbsoluteUrl = String(spi.webAbsoluteUrl || "").replace(/\/$/, "");
  try {
    var u = new URL(window.location.href);
    var vp =
      u.searchParams.get("view") ||
      u.searchParams.get("View") ||
      u.searchParams.get("viewid") ||
      u.searchParams.get("ViewId");
    if (vp) viewId = normGuid(decodeURIComponent(vp));
    var lp = u.searchParams.get("list") || u.searchParams.get("List");
    if (lp) listId = normGuid(decodeURIComponent(lp));
  } catch (e) {}
  window.postMessage(
    {
      __spcsv: true,
      type: "SPCSVViewFormatContext",
      listId: listId,
      viewId: viewId,
      webAbsoluteUrl: webAbsoluteUrl,
    },
    "*"
  );
})();
