/**
 * SharePoint ViewFields planning helpers.
 * Loaded before views.js; also runnable under Node tests (module.exports).
 */
(function (root) {
  "use strict";

  function viewFieldsResponseToNames(vfJson) {
    var d = vfJson && vfJson.d;
    var vfItems = (vfJson && (vfJson.Items || vfJson.value || (d && d.results))) || [];
    return []
      .concat(vfItems)
      .filter(Boolean)
      .map(function (x) {
        return String(typeof x === "object" ? x.Name || x.name || x : x);
      })
      .filter(Boolean);
  }

  function uniqueNames(names) {
    var seen = Object.create(null);
    var out = [];
    for (var i = 0; i < (names || []).length; i++) {
      var name = String(names[i] || "").trim();
      if (!name || seen[name]) continue;
      seen[name] = true;
      out.push(name);
    }
    return out;
  }

  function planViewFieldChanges(currentNames, desiredNames) {
    var current = uniqueNames(currentNames);
    var desired = uniqueNames(desiredNames);
    var currentSet = Object.create(null);
    var desiredSet = Object.create(null);
    var additions = [];
    var removals = [];
    var moves = [];
    var i;

    for (i = 0; i < current.length; i++) currentSet[current[i]] = true;
    for (i = 0; i < desired.length; i++) {
      desiredSet[desired[i]] = true;
      if (!currentSet[desired[i]]) additions.push(desired[i]);
      moves.push({ field: desired[i], index: i });
    }
    for (i = 0; i < current.length; i++) {
      if (!desiredSet[current[i]]) removals.push(current[i]);
    }

    return { additions: additions, removals: removals, moves: moves };
  }

  var api = {
    viewFieldsResponseToNames: viewFieldsResponseToNames,
    uniqueNames: uniqueNames,
    planViewFieldChanges: planViewFieldChanges
  };

  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }
  if (typeof window !== "undefined") {
    window.SPViewFieldsCore = api;
  }
})(typeof globalThis !== "undefined" ? globalThis : this);
