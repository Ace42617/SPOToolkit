/* global window */
(function () {
  function cleanListId(value) {
    return String(value || "").replace(/[{}]/g, "").trim();
  }

  function cleanSiteUrl(value) {
    return String(value || "").replace(/\/$/, "").trim();
  }

  function resolveCompassActionContext(invoke, action, requireList) {
    if (typeof invoke !== "function") {
      return Promise.reject(new Error("Could not resolve the current SharePoint page."));
    }
    return Promise.resolve()
      .then(function () {
        return invoke({ action: action || "getPageContext" });
      })
      .then(function (response) {
        if (!response || response.ok === false) {
          throw new Error((response && response.error) || "Could not resolve the current SharePoint page.");
        }
        const siteUrl = cleanSiteUrl(
          response.webAbsoluteUrl || response.siteAbsoluteUrl || response.siteUrl
        );
        const listId = cleanListId(response.pageListId || response.listId);
        if (!siteUrl || (requireList && !listId)) {
          throw new Error(
            requireList
              ? "Open a list or library view before running this action."
              : "Open a SharePoint site before running this action."
          );
        }
        return Object.assign({}, response, { siteUrl: siteUrl, listId: listId });
      });
  }

  window.SPOT_resolveCompassActionContext = resolveCompassActionContext;
})();
