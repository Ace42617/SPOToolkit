/* Restore SharePoint's define and load the export script. URL is in this script's query param "url". */
(function () {
  if (typeof window === "undefined") return;
  if (typeof window.__spcsv_define === "function") {
    window.define = window.__spcsv_define;
    if (window.__spcsv_define_amd !== undefined) {
      window.define.amd = window.__spcsv_define_amd;
    }
    try {
      delete window.__spcsv_define;
      delete window.__spcsv_define_amd;
    } catch (_) {
      window.__spcsv_define = undefined;
      window.__spcsv_define_amd = undefined;
    }
  }
  var script = document.currentScript;
  var url = "";
  if (script && script.src) {
    try {
      var u = new URL(script.src);
      var p = u.searchParams.get("url");
      if (p) url = decodeURIComponent(p);
    } catch (e) {}
  }
  if (url) {
    var s = document.createElement("script");
    s.src = url;
    (document.head || document.documentElement).appendChild(s);
  }
})();
