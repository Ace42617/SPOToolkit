import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

async function loadCommonJsExport(filePath) {
  const source = await readFile(filePath, "utf8");
  const mod = { exports: {} };
  const load = new Function("module", "exports", source + "\nreturn module.exports;");
  return load(mod, mod.exports);
}

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const { kindsFromTypeAsString, suggest } = await loadCommonJsExport(path.join(repoRoot, "filterTypeaheadLogic.js"));

describe("filter typeahead logic", () => {
  it("maps SharePoint field types to suggestion kinds", () => {
    assert.deepEqual(kindsFromTypeAsString("DateTime"), ["date"]);
    assert.deepEqual(kindsFromTypeAsString("date"), ["date"]);
    assert.deepEqual(kindsFromTypeAsString("User"), ["person"]);
    assert.deepEqual(kindsFromTypeAsString("UserMulti"), ["person"]);
    assert.deepEqual(kindsFromTypeAsString("Text"), []);
    assert.deepEqual(kindsFromTypeAsString("Number"), []);
    assert.deepEqual(kindsFromTypeAsString("Lookup"), []);
  });

  it("suggests date tokens for date fields", () => {
    const d = kindsFromTypeAsString("DateTime");
    const dateEmpty = suggest(d, "");
    assert.ok(dateEmpty.length > 0);
    assert.ok(dateEmpty.every((x) => x.kind === "date"));
    assert.ok(dateEmpty[0].token.indexOf("[Today]") === 0);
  });

  it("suggests person tokens for person fields", () => {
    const p = kindsFromTypeAsString("User");
    const personEmpty = suggest(p, "");
    assert.equal(personEmpty.length, 1);
    assert.equal(personEmpty[0].token, "[Me]");
  });

  it("does not suggest tokens for unsupported field types", () => {
    assert.deepEqual(suggest([], ""), []);
  });

  it("keeps date and person-only suggestions separated", () => {
    const d = kindsFromTypeAsString("DateTime");
    const dateMe = suggest(d, "me");
    assert.ok(dateMe.every((x) => x.token.indexOf("[Me]") < 0), "date column should not suggest [Me]");

    const p = kindsFromTypeAsString("User");
    const personWeek = suggest(p, "week");
    assert.equal(personWeek.length, 0, "person column should not suggest date-only tokens for 'week'");
  });
});
