import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import test from "node:test";
import { cleanSharePointDelimitedValue, flattenRestValue } from "../lib/exportValueClean.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");

function extractNamedFunction(source, name) {
  const needle = "function " + name + "(";
  const start = source.indexOf(needle);
  if (start < 0) throw new Error("function " + name + " not found");
  const brace = source.indexOf("{", start);
  if (brace < 0) throw new Error("function " + name + " has no body");
  let depth = 0;
  for (let j = brace; j < source.length; j++) {
    const ch = source[j];
    if (ch === "{") depth++;
    else if (ch === "}") {
      depth--;
      if (depth === 0) return source.slice(start, j + 1);
    }
  }
  throw new Error("function " + name + " is unbalanced");
}

const exportCsv = readFileSync(join(root, "exportCSV.js"), "utf8");
const cleanValue = new Function("return (" + extractNamedFunction(exportCsv, "cleanValue") + ")")();
const flattenFromCsv = new Function("return (" + extractNamedFunction(exportCsv, "flattenRestValue") + ")")();

const CASES = [
  ["plain text", "Hello", "Hello"],
  ["empty", "", ""],
  ["lookup single", "12;#Design", "Design"],
  ["lookup multi", "12;#Design;#34;#Engineering", "Design, Engineering"],
  ["person", "23;#Jane Doe", "Jane Doe"],
  ["multi-choice owssvr two", ";#Design;#Engineering;#", "Design, Engineering"],
  ["multi-choice owssvr three", ";#Design;#Engineering;#Legal;#", "Design, Engineering, Legal"],
  ["multi-choice owssvr one", ";#Design;#", "Design"],
  ["REST-flattened multi-choice", "Design;#Engineering;#Legal", "Design, Engineering, Legal"],
  ["REST-flattened two", "Red;#Blue", "Red, Blue"],
  ["no delimiter numeric", "42", "42"],
  ["text containing ;# but not lookup ids", "see item 12;#urgent", "see item 12, urgent"],
];

test("cleanSharePointDelimitedValue keeps every MultiChoice value", () => {
  for (const [label, input, expected] of CASES) {
    assert.equal(cleanSharePointDelimitedValue(input), expected, label);
  }
  assert.equal(cleanSharePointDelimitedValue(null), "");
  assert.equal(cleanSharePointDelimitedValue(undefined), "");
});

test("exportCSV.js cleanValue matches helper (source-sync)", () => {
  for (const [label, input, expected] of CASES) {
    assert.equal(cleanValue(input), expected, "exportCSV " + label);
    assert.equal(cleanValue(input), cleanSharePointDelimitedValue(input), "parity " + label);
  }
});

test("old odd-index lookup parse dropped MultiChoice values", () => {
  const broken = (s) => {
    const parts = String(s).split(";#");
    return parts.filter(function (_, i) { return i % 2 === 1; }).join(", ") || s;
  };
  assert.equal(broken(";#Design;#Engineering;#Legal;#"), "Design, Legal");
  assert.equal(broken("Design;#Engineering;#Legal"), "Engineering");
  assert.notEqual(cleanSharePointDelimitedValue(";#Design;#Engineering;#Legal;#"), broken(";#Design;#Engineering;#Legal;#"));
  assert.equal(cleanSharePointDelimitedValue(";#Design;#Engineering;#Legal;#"), "Design, Engineering, Legal");
});

test("flattenRestValue joins MultiChoice arrays with commas, not ;#", () => {
  assert.equal(flattenRestValue(["Design", "Engineering", "Legal"]), "Design, Engineering, Legal");
  assert.equal(flattenRestValue(["Red", "Blue"]), "Red, Blue");
  assert.equal(flattenRestValue([{ LookupValue: "A" }, { LookupValue: "B" }]), "A, B");
  assert.equal(flattenRestValue({ LookupValue: "Solo" }), "Solo");
  assert.equal(flattenRestValue("plain"), "plain");
  assert.equal(flattenFromCsv(["Design", "Engineering", "Legal"]), "Design, Engineering, Legal");
  assert.equal(flattenFromCsv(["Design", "Engineering"]), flattenRestValue(["Design", "Engineering"]));
  assert.ok(!flattenFromCsv(["Design", "Engineering"]).includes(";#"));
});
