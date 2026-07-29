import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import {
  formatRestVersionsFetchFailureMessage,
  isRestVersionsFetchFailure
} from "../lib/restVersionFetchIntegrity.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const exportSources = [
  "exportCSV.js",
  "SP-Developer-Toolkit-Experimental/exportCSV.js",
  "SP-Developer-Toolkit-Lite/exportCSV.js"
].map((rel) => ({
  rel,
  source: readFileSync(join(root, rel), "utf8")
}));

describe("isRestVersionsFetchFailure", () => {
  it("treats null/undefined as failure", () => {
    assert.equal(isRestVersionsFetchFailure(null), true);
    assert.equal(isRestVersionsFetchFailure(undefined), true);
  });

  it("does not treat empty or populated arrays as failure", () => {
    assert.equal(isRestVersionsFetchFailure([]), false);
    assert.equal(isRestVersionsFetchFailure([{ VersionLabel: "1.0" }]), false);
  });
});

describe("formatRestVersionsFetchFailureMessage", () => {
  it("includes the item id and says no file was downloaded", () => {
    const msg = formatRestVersionsFetchFailureMessage(42);
    assert.match(msg, /item 42/);
    assert.match(msg, /No file downloaded/i);
  });
});

describe("exportCSV.js variants fail closed on REST versions fetch failure", () => {
  for (const { rel, source } of exportSources) {
    it(`${rel} rejects null versions results before current-item fallthrough`, () => {
      assert.match(source, /function isRestVersionsFetchFailure\s*\(/);
      assert.match(source, /function formatRestVersionsFetchFailureMessage\s*\(/);
      assert.match(source, /isRestVersionsFetchFailure\s*\(\s*versions\s*\)/);
      assert.match(source, /formatRestVersionsFetchFailureMessage\s*\(\s*itemId\s*\)/);
      assert.match(source, /No file downloaded/);
      // Former silent truncation: null and [] both hit the current-item else branch.
      assert.doesNotMatch(
        source,
        /var versions = await fetchAllVersionsForItem\([^)]+\);\s*if \(versions && versions\.length > 0\)/
      );
    });
  }
});
