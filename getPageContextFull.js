(function () {
  var i = window._spPageContextInfo || null;
  var data = null;
  try {
    data = i ? JSON.parse(JSON.stringify(i)) : null;
  } catch (_) {
    data = i || null;
  }
  window.postMessage(
    {
      __spcsv: true,
      type: "SPCSVPageContextFull",
      data: data
    },
    "*"
  );
})();
