import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { evaluateViewFieldsBatchResult } from "../lib/viewFieldsBatchIntegrity.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const rootExport = readFileSync(join(root, "exportCSV.js"), "utf8");
const experimentalExport = readFileSync(
  join(root, "SP-Developer-Toolkit-Experimental/exportCSV.js"),
  "utf8"
);

describe("evaluateViewFieldsBatchResult", () => {
  it("accepts a fully successful batch", () => {
    assert.deepEqual(evaluateViewFieldsBatchResult(0, 12), { ok: true });
  });

  it("fails closed when any column fails to attach", () => {
    const oneOfTwenty = evaluateViewFieldsBatchResult(1, 20);
    assert.equal(oneOfTwenty.ok, false);
    assert.match(oneOfTwenty.error, /1 of 20/);
    assert.match(oneOfTwenty.error, /No file downloaded/i);

    // Former half-failure tolerance allowed exactly half to fail.
    const half = evaluateViewFieldsBatchResult(5, 10);
    assert.equal(half.ok, false);
    assert.match(half.error, /5 of 10/);
  });

  it("fails closed when there are no columns to attach", () => {
    const empty = evaluateViewFieldsBatchResult(0, 0);
    assert.equal(empty.ok, false);
    assert.match(empty.error, /no columns to attach/i);
    assert.match(empty.error, /No file downloaded/i);
  });
});

describe("exportCSV.js setViewFieldsBatch stays fail-closed", () => {
  for (const [label, source] of [
    ["root", rootExport],
    ["Experimental", experimentalExport]
  ]) {
    it(`${label}: rejects any failed addViewField (no half-failure tolerance)`, () => {
      assert.match(source, /async function setViewFieldsBatch\s*\(/);
      assert.match(
        source,
        /View field setup failed for "\s*\+\s*failed\s*\+\s*" of "\s*\+\s*names\.length/
      );
      assert.match(source, /No file downloaded/);
      // Former silent truncation gate.
      assert.doesNotMatch(source, /failed\s*>\s*names\.length\s*\/\s*2/);
    });
  }
});
