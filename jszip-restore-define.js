(function () {
  if (typeof window === "undefined") return;
  if (typeof window.__spcsv_define !== "function") return;
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
})();
