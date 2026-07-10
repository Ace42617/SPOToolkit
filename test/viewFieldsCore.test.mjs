import { readFileSync } from "node:fs";
import assert from "node:assert/strict";

function loadBrowserGlobalScript(path) {
  const source = readFileSync(new URL(path, import.meta.url), "utf8");
  const module = { exports: {} };
  const window = {};
  Function("module", "exports", "window", source)(module, module.exports, window);
  return { exports: module.exports, window };
}

const { exports: core, window } = loadBrowserGlobalScript("../viewFieldsCore.js");

assert.equal(window.SPViewFieldsCore, core);

assert.deepEqual(core.viewFieldsResponseToNames({ Items: ["Title", { Name: "Modified" }] }), ["Title", "Modified"]);
assert.deepEqual(core.viewFieldsResponseToNames({ d: { results: [{ name: "LinkTitle" }] } }), ["LinkTitle"]);
assert.deepEqual(core.uniqueNames(["Title", "Title", "", "Modified"]), ["Title", "Modified"]);

const addPlan = core.planViewFieldChanges(["Title", "Modified"], ["Title", "Status", "Modified"]);
assert.deepEqual(addPlan.additions, ["Status"]);
assert.deepEqual(addPlan.removals, []);
assert.deepEqual(addPlan.moves, [
  { field: "Title", index: 0 },
  { field: "Status", index: 1 },
  { field: "Modified", index: 2 },
]);

const removePlan = core.planViewFieldChanges(["Title", "Obsolete", "Modified"], ["Modified", "Title"]);
assert.deepEqual(removePlan.additions, []);
assert.deepEqual(removePlan.removals, ["Obsolete"]);
assert.deepEqual(removePlan.moves, [
  { field: "Modified", index: 0 },
  { field: "Title", index: 1 },
]);

for (const path of ["../views.js", "../SP-Developer-Toolkit-Experimental/views.js"]) {
  const source = readFileSync(new URL(path, import.meta.url), "utf8");
  assert.ok(!source.includes("RemoveAllViewFields"), `${path} must not clear view fields before saving`);
  assert.match(source, /moveViewFieldTo/);
  assert.match(source, /removeviewfield/);
  assert.match(source, /viewDetailsLoadedForId !== selectedViewId/);
}

console.log("viewFieldsCore.test.mjs: OK");
