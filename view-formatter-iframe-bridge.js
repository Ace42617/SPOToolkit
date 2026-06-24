// Runs in SharePoint documents loaded inside the View Formatter preview iframe only.
// Inspired by pnp/sp-formatter's React-fiber integration + postMessage; list pages have no format-pane Preview button.
(function () {
  if (window.self === window.top) return;

  var bridge = false;
  try {
    bridge = new URLSearchParams(window.location.search || "").get("_spotkVf") === "1";
  } catch (_) {}
  if (!bridge) {
    try {
      bridge = (document.referrer || "").indexOf("view-formatter.html") >= 0;
    } catch (_) {}
  }
  if (!bridge) return;

  if (window.__spotkVfBridgeInstalled) return;
  window.__spotkVfBridgeInstalled = true;

  var lastExtOrigin = "";

  function fiberKey(el) {
    if (!el || typeof el !== "object") return null;
    for (var k in el) {
      if (
        Object.prototype.hasOwnProperty.call(el, k) &&
        (k.indexOf("__reactFiber$") === 0 || k.indexOf("__reactInternalInstance$") === 0)
      ) {
        return k;
      }
    }
    return null;
  }

  function getFiber(el) {
    var fk = fiberKey(el);
    return fk ? el[fk] : null;
  }

  function walkFibers(f, visit) {
    if (!f) return;
    visit(f);
    walkFibers(f.child, visit);
    walkFibers(f.sibling, visit);
  }

  function tryForceUpdateFrom(fiberStart) {
    var f = fiberStart;
    var depth = 0;
    while (f && depth < 80) {
      try {
        var sn = f.stateNode;
        if (sn && typeof sn.forceUpdate === "function") sn.forceUpdate();
      } catch (_) {}
      f = f.return;
      depth++;
    }
  }

  function considerPatch(obj, fiber, compact, patchKeys, patchedHosts) {
    if (!obj || typeof obj !== "object") return 0;
    var n = 0;
    for (var i = 0; i < patchKeys.length; i++) {
      var pk = patchKeys[i];
      if (!Object.prototype.hasOwnProperty.call(obj, pk)) continue;
      var cur = obj[pk];
      if (cur !== undefined && cur !== null && typeof cur !== "string") continue;
      try {
        obj[pk] = compact;
        n++;
        if (fiber && patchedHosts.indexOf(fiber) === -1) patchedHosts.push(fiber);
      } catch (_) {}
    }
    if (obj.viewFields && typeof obj.viewFields === "object" && obj.customFormatter === undefined) {
      try {
        obj.customFormatter = compact;
        n++;
        if (fiber && patchedHosts.indexOf(fiber) === -1) patchedHosts.push(fiber);
      } catch (_) {}
    }
    return n;
  }

  function applyViewFormatJson(jsonStr) {
    var compact;
    try {
      compact = JSON.stringify(JSON.parse(typeof jsonStr === "string" ? jsonStr : "{}"));
    } catch (_) {
      return { ok: false, step: "bad-json" };
    }

    var patchKeys = [
      "customFormatter",
      "CustomFormatter",
      "viewFormatter",
      "viewFormatting",
      "ViewFormatting",
      "listViewFormatting",
      "formatterDefinition",
      "listFormatting",
      "formattingDefinition",
      "viewFormat",
      "rowFormatterDefinition",
    ];

    var rootCandidates = [];
    var seen = Object.create(null);
    function addRoot(el) {
      if (!el || seen[el]) return;
      seen[el] = 1;
      rootCandidates.push(el);
    }

    addRoot(document.getElementById("spPageCanvasContent"));
    addRoot(document.getElementById("spPageContainerContent"));
    addRoot(document.getElementById("workbenchPageContent"));
    try {
      document.querySelectorAll("[data-sp-feature-tag='List'],[data-automationid='List']").forEach(function (el) {
        addRoot(el);
      });
    } catch (_) {}
    addRoot(document.body);

    var patchedHosts = [];
    var patched = 0;
    for (var r = 0; r < rootCandidates.length; r++) {
      var rf = getFiber(rootCandidates[r]);
      if (!rf) continue;
      walkFibers(rf, function (fib) {
        var mp = fib.memoizedProps;
        var pp = fib.pendingProps;
        if (mp) {
          patched += considerPatch(mp, fib, compact, patchKeys, patchedHosts);
          if (mp.props) patched += considerPatch(mp.props, fib, compact, patchKeys, patchedHosts);
        }
        if (pp && pp !== mp) {
          patched += considerPatch(pp, fib, compact, patchKeys, patchedHosts);
          if (pp.props) patched += considerPatch(pp.props, fib, compact, patchKeys, patchedHosts);
        }
      });
    }

    if (!patched) return { ok: false, step: "no-target" };

    function rerender() {
      for (var h = 0; h < patchedHosts.length; h++) {
        tryForceUpdateFrom(patchedHosts[h]);
      }
      var b = getFiber(document.body);
      if (b) tryForceUpdateFrom(b);
    }
    rerender();
    try {
      requestAnimationFrame(rerender);
    } catch (_) {}

    return { ok: true, patched: patched };
  }

  window.addEventListener(
    "message",
    function (ev) {
      if (typeof ev.origin !== "string") return;
      if (ev.origin.indexOf("chrome-extension://") !== 0) return;
      lastExtOrigin = ev.origin;
      if (!ev.data || ev.data.__spotkVfPreview !== true) return;
      var res = applyViewFormatJson(ev.data.json);
      try {
        if (ev.source && lastExtOrigin) {
          ev.source.postMessage({ __spotkVfPreviewResult: true, result: res }, lastExtOrigin);
        }
      } catch (_) {}
    },
    false
  );
})();
