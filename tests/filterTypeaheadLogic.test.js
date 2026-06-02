"use strict";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const __dirname = dirname(fileURLToPath(import.meta.url));
const source = await readFile(join(__dirname, "../filterTypeaheadLogic.js"), "utf8");
const module = { exports: {} };
Function("module", "window", source)(module, {});
const { kindsFromTypeAsString, suggest } = module.exports;

assert.deepStrictEqual(kindsFromTypeAsString("DateTime"), ["date"]);
assert.deepStrictEqual(kindsFromTypeAsString("date"), ["date"]);
assert.deepStrictEqual(kindsFromTypeAsString("User"), ["person"]);
assert.deepStrictEqual(kindsFromTypeAsString("UserMulti"), ["person"]);
assert.deepStrictEqual(kindsFromTypeAsString("Text"), []);
assert.deepStrictEqual(kindsFromTypeAsString("Number"), []);
assert.deepStrictEqual(kindsFromTypeAsString("Lookup"), []);

const d = kindsFromTypeAsString("DateTime");
const dateEmpty = suggest(d, "");
assert.ok(dateEmpty.length > 0);
assert.ok(dateEmpty.every((x) => x.kind === "date"));
assert.ok(dateEmpty[0].token.indexOf("[Today]") === 0);

const p = kindsFromTypeAsString("User");
const personEmpty = suggest(p, "");
assert.strictEqual(personEmpty.length, 1);
assert.strictEqual(personEmpty[0].token, "[Me]");

assert.deepStrictEqual(suggest([], ""), []);

const dateMe = suggest(d, "me");
assert.ok(dateMe.every((x) => x.token.indexOf("[Me]") < 0), "date column should not suggest [Me]");

const personWeek = suggest(p, "week");
assert.strictEqual(personWeek.length, 0, "person column should not suggest date-only tokens for 'week'");

console.log("filterTypeaheadLogic.test.js: OK");
