/* Hide AMD/define so JSZip attaches to window.JSZip on pages that have RequireJS (e.g. SharePoint). */
(function () {
  if (typeof window === "undefined") return;
  window.__spcsv_define = window.define;
  try { window.define = undefined; } catch (e) {}
})();
