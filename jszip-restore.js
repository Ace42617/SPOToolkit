/* Restore define and load the export script. URL is in this script's query param "url". */
(function () {
  if (typeof window === "undefined") return;
  try { window.define = window.__spcsv_define; } catch (e) {}
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
