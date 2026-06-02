(function (root) {
  "use strict";

  function uniqueFieldNames(names) {
    var out = [];
    var seen = Object.create(null);
    for (var i = 0; i < (names || []).length; i++) {
      var name = String(names[i] == null ? "" : names[i]).trim();
      if (!name || seen[name]) continue;
      seen[name] = true;
      out.push(name);
    }
    return out;
  }

  function viewFieldsResponseToNames(data) {
    var d = data && data.d;
    var items = (data && (data.Items || data.value || (d && d.results))) || [];
    var out = [];
    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      if (!item) continue;
      out.push(String(typeof item === "object" ? item.Name || item : item));
    }
    return uniqueFieldNames(out);
  }

  function asSet(names) {
    var set = Object.create(null);
    for (var i = 0; i < names.length; i++) set[names[i]] = true;
    return set;
  }

  function buildViewFieldUpdatePlan(currentFields, desiredFields) {
    var current = uniqueFieldNames(currentFields);
    var desired = uniqueFieldNames(desiredFields);
    var currentSet = asSet(current);
    var desiredSet = asSet(desired);
    var additions = [];
    var removals = [];
    var moves = [];
    var i;

    for (i = 0; i < desired.length; i++) {
      if (!currentSet[desired[i]]) additions.push(desired[i]);
      moves.push({ field: desired[i], index: i });
    }
    for (i = 0; i < current.length; i++) {
      if (!desiredSet[current[i]]) removals.push(current[i]);
    }

    return {
      current: current,
      desired: desired,
      additions: additions,
      removals: removals,
      moves: moves
    };
  }

  async function applyViewFieldUpdatePlan(plan, ops) {
    for (var i = 0; i < plan.additions.length; i++) {
      await ops.add(plan.additions[i]);
    }
    for (var j = 0; j < plan.removals.length; j++) {
      await ops.remove(plan.removals[j]);
    }
    for (var k = 0; k < plan.moves.length; k++) {
      await ops.move(plan.moves[k].field, plan.moves[k].index);
    }
  }

  var api = {
    uniqueFieldNames: uniqueFieldNames,
    viewFieldsResponseToNames: viewFieldsResponseToNames,
    buildViewFieldUpdatePlan: buildViewFieldUpdatePlan,
    applyViewFieldUpdatePlan: applyViewFieldUpdatePlan
  };

  if (typeof module !== "undefined" && module.exports) module.exports = api;
  root.SPOToolkitViewFieldsCore = api;
})(typeof window !== "undefined" ? window : (typeof globalThis !== "undefined" ? globalThis : this));
