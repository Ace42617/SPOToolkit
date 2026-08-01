"use strict";
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const code = fs.readFileSync(path.join(__dirname, "..", "filterTypeaheadLogic.js"), "utf8");
const sandbox = { module: { exports: {} }, globalThis: {} };
sandbox.exports = sandbox.module.exports;
vm.runInNewContext(code, sandbox);
const api = sandbox.module.exports;

function kinds(type) {
  return Array.from(api.kindsFromTypeAsString(type));
}
function suggest(kindsArr, q) {
  return api.suggest(kindsArr, q).map((x) => ({
    kind: x.kind,
    token: x.token,
  }));
}

assert.deepStrictEqual(kinds("DateTime"), ["date"]);
assert.deepStrictEqual(kinds("date"), ["date"]);
assert.deepStrictEqual(kinds("User"), ["person"]);
assert.deepStrictEqual(kinds("UserMulti"), ["person"]);
assert.deepStrictEqual(kinds("Text"), []);
assert.deepStrictEqual(kinds("Number"), []);
assert.deepStrictEqual(kinds("Lookup"), []);

const d = kinds("DateTime");
const dateEmpty = suggest(d, "");
assert.ok(dateEmpty.length > 0);
assert.ok(dateEmpty.every((x) => x.kind === "date"));
assert.ok(dateEmpty[0].token.indexOf("[Today]") === 0);

const p = kinds("User");
const personEmpty = suggest(p, "");
assert.strictEqual(personEmpty.length, 1);
assert.strictEqual(personEmpty[0].token, "[Me]");

assert.deepStrictEqual(suggest([], ""), []);

const dateMe = suggest(d, "me");
assert.ok(dateMe.every((x) => x.token.indexOf("[Me]") < 0), "date column should not suggest [Me]");

const personWeek = suggest(p, "week");
assert.strictEqual(personWeek.length, 0, "person column should not suggest date-only tokens for 'week'");
