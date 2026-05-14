import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import vm from "node:vm";

const moduleShim = { exports: {} };
const source = readFileSync(new URL("../filterTypeaheadLogic.js", import.meta.url), "utf8");
vm.runInNewContext(
  source,
  {
    module: moduleShim,
    exports: moduleShim.exports,
    console,
  },
  { filename: "filterTypeaheadLogic.js" }
);
const { kindsFromTypeAsString, suggest } = moduleShim.exports;

assert.deepStrictEqual(Array.from(kindsFromTypeAsString("DateTime")), ["date"]);
assert.deepStrictEqual(Array.from(kindsFromTypeAsString("date")), ["date"]);
assert.deepStrictEqual(Array.from(kindsFromTypeAsString("User")), ["person"]);
assert.deepStrictEqual(Array.from(kindsFromTypeAsString("UserMulti")), ["person"]);
assert.deepStrictEqual(Array.from(kindsFromTypeAsString("Text")), []);
assert.deepStrictEqual(Array.from(kindsFromTypeAsString("Number")), []);
assert.deepStrictEqual(Array.from(kindsFromTypeAsString("Lookup")), []);

const d = kindsFromTypeAsString("DateTime");
const dateEmpty = suggest(d, "");
assert.ok(dateEmpty.length > 0);
assert.ok(dateEmpty.every((x) => x.kind === "date"));
assert.ok(dateEmpty[0].token.indexOf("[Today]") === 0);

const p = kindsFromTypeAsString("User");
const personEmpty = suggest(p, "");
assert.strictEqual(personEmpty.length, 1);
assert.strictEqual(personEmpty[0].token, "[Me]");

assert.deepStrictEqual(Array.from(suggest([], "")), []);

const dateMe = suggest(d, "me");
assert.ok(dateMe.every((x) => x.token.indexOf("[Me]") < 0), "date column should not suggest [Me]");

const personWeek = suggest(p, "week");
assert.strictEqual(personWeek.length, 0, "person column should not suggest date-only tokens for 'week'");

console.log("filterTypeaheadLogic.test.js: OK");
