/* Keep SharePoint's AMD loader working while JSZip loads on the global branch. */
(function () {
  if (typeof window === "undefined") return;
  var realDefine = window.define;
  if (typeof realDefine !== "function") return;
  window.__spcsv_define = realDefine;
  window.__spcsv_define_amd = realDefine.amd;
  window.define = function () {
    return realDefine.apply(this, arguments);
  };
  window.define.amd = false;
})();
