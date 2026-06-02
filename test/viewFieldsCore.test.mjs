/**
 * Tests for safe View Manager ViewFields update planning.
 */
import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const __dirname = dirname(fileURLToPath(import.meta.url));

async function loadViewFieldsCore() {
  const source = await readFile(join(__dirname, "../viewFieldsCore.js"), "utf8");
  const module = { exports: {} };
  Function("module", "window", source)(module, {});
  return module.exports;
}

describe("ViewFields update plan", () => {
  it("adds missing selected fields before removing or moving existing fields", async () => {
    const core = await loadViewFieldsCore();
    const plan = core.buildViewFieldUpdatePlan(
      ["Title", "Modified", "Obsolete"],
      ["Modified", "Title", "Status"]
    );

    assert.deepEqual(plan.additions, ["Status"]);
    assert.deepEqual(plan.removals, ["Obsolete"]);
    assert.deepEqual(plan.moves, [
      { field: "Modified", index: 0 },
      { field: "Title", index: 1 },
      { field: "Status", index: 2 },
    ]);

    const calls = [];
    await core.applyViewFieldUpdatePlan(plan, {
      add: async (name) => calls.push("add:" + name),
      remove: async (name) => calls.push("remove:" + name),
      move: async (name, index) => calls.push("move:" + name + ":" + index),
    });

    assert.deepEqual(calls, [
      "add:Status",
      "remove:Obsolete",
      "move:Modified:0",
      "move:Title:1",
      "move:Status:2",
    ]);
  });

  it("does not remove existing fields when a desired field cannot be added", async () => {
    const core = await loadViewFieldsCore();
    const plan = core.buildViewFieldUpdatePlan(["Title", "Obsolete"], ["Title", "Missing"]);
    const calls = [];

    await assert.rejects(
      () =>
        core.applyViewFieldUpdatePlan(plan, {
          add: async (name) => {
            calls.push("add:" + name);
            throw new Error("add failed");
          },
          remove: async (name) => calls.push("remove:" + name),
          move: async (name, index) => calls.push("move:" + name + ":" + index),
        }),
      /add failed/
    );

    assert.deepEqual(calls, ["add:Missing"]);
  });

  it("normalizes SharePoint ViewFields payload shapes", async () => {
    const core = await loadViewFieldsCore();
    assert.deepEqual(core.viewFieldsResponseToNames({ Items: ["Title", { Name: "Modified" }, "Title"] }), [
      "Title",
      "Modified",
    ]);
    assert.deepEqual(core.viewFieldsResponseToNames({ d: { results: [{ Name: "LinkTitle" }] } }), ["LinkTitle"]);
  });
});
