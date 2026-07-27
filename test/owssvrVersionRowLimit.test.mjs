import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import {
  computeOwssvrRowLimit,
  formatOwssvrVersionTruncationMessage,
  isOwssvrVersionPageTruncated,
  versionedDocLibIdPageLimit
} from "../lib/owssvrVersionRowLimit.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const exportSources = [
  "exportCSV.js",
  "SP-Developer-Toolkit-Experimental/exportCSV.js",
  "SP-Developer-Toolkit-Lite/exportCSV.js"
].map((rel) => ({
  rel,
  source: readFileSync(join(root, rel), "utf8")
}));

describe("computeOwssvrRowLimit", () => {
  it("uses RowLimit=0 for versioned lists", () => {
    assert.equal(computeOwssvrRowLimit(1000, true, false), 0);
  });

  it("keeps a finite RowLimit for versioned document libraries", () => {
    assert.equal(computeOwssvrRowLimit(1000, true, true), 50000);
    assert.equal(computeOwssvrRowLimit(5000, true, true), 100000);
  });

  it("keeps a finite RowLimit when versions are off", () => {
    assert.equal(computeOwssvrRowLimit(500, false, true), 50000);
    assert.equal(computeOwssvrRowLimit(500, false, false), 50000);
  });
});

describe("versionedDocLibIdPageLimit", () => {
  it("shrinks large page sizes so version expansion has headroom", () => {
    assert.equal(versionedDocLibIdPageLimit(5000, 50000), 500);
    assert.equal(versionedDocLibIdPageLimit(1000, 50000), 500);
    assert.equal(versionedDocLibIdPageLimit(200, 50000), 200);
  });
});

describe("isOwssvrVersionPageTruncated / formatOwssvrVersionTruncationMessage", () => {
  it("flags saturation only for versioned document libraries", () => {
    assert.equal(
      isOwssvrVersionPageTruncated({
        includeVersions: true,
        isDocLib: true,
        rowCount: 50000,
        rowLimit: 50000
      }),
      true
    );
    assert.equal(
      isOwssvrVersionPageTruncated({
        includeVersions: true,
        isDocLib: true,
        rowCount: 49999,
        rowLimit: 50000
      }),
      false
    );
    assert.equal(
      isOwssvrVersionPageTruncated({
        includeVersions: false,
        isDocLib: true,
        rowCount: 50000,
        rowLimit: 50000
      }),
      false
    );
    assert.equal(
      isOwssvrVersionPageTruncated({
        includeVersions: true,
        isDocLib: false,
        rowCount: 50000,
        rowLimit: 0
      }),
      false
    );
  });

  it("tells the operator that no file was downloaded", () => {
    const msg = formatOwssvrVersionTruncationMessage(50000, 5000);
    assert.match(msg, /RowLimit \(50000 rows\)/);
    assert.match(msg, /ID page of 5000/);
    assert.match(msg, /No file downloaded/i);
  });
});

describe("exportCSV.js variants fail closed on version RowLimit saturation", () => {
  for (const { rel, source } of exportSources) {
    it(`${rel} detects OWSSVR version RowLimit truncation and does not download`, () => {
      assert.match(source, /function computeOwssvrRowLimit\s*\(/);
      assert.match(source, /function isOwssvrVersionPageTruncated\s*\(/);
      assert.match(source, /function formatOwssvrVersionTruncationMessage\s*\(/);
      assert.match(source, /No file downloaded/);
      assert.match(
        source,
        /isOwssvrVersionPageTruncated\s*\(\s*\{[\s\S]*?rowCount:\s*parsed\.rows\.length/
      );
      // Former silent success path must not treat saturation as a normal page.
      assert.doesNotMatch(
        source,
        /var rowLimit = \(incVer && !isDocLib\) \? 0 : Math\.max\(PAGE_LIMIT \* 20, 50000\);/
      );
    });
  }

  it("root and Experimental shrink versioned doc-lib ID pages", () => {
    for (const rel of ["exportCSV.js", "SP-Developer-Toolkit-Experimental/exportCSV.js"]) {
      const source = exportSources.find((e) => e.rel === rel).source;
      assert.match(source, /function versionedDocLibIdPageLimit\s*\(/);
      assert.match(source, /idPageLimit\s*=\s*versionedDocLibIdPageLimit\s*\(/);
    }
  });
});
