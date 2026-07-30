import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import {
  canReportExportCsvMatrixSuccess,
  formatMatrixItemScanFailureMessage,
  formatMatrixRoleDefsEmptyMessage,
  formatMatrixRoleDefsLoadFailureMessage
} from "../lib/exportCsvMatrixIntegrity.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const exportSources = [
  "exportCSV.js",
  "SP-Developer-Toolkit-Experimental/exportCSV.js",
  "SP-Developer-Toolkit-Lite/exportCSV.js"
].map((rel) => ({
  rel,
  source: readFileSync(join(root, rel), "utf8")
}));

describe("formatMatrixRoleDefsLoadFailureMessage / empty", () => {
  it("names the failure and says no file was downloaded", () => {
    assert.match(formatMatrixRoleDefsLoadFailureMessage("HTTP 403"), /HTTP 403/);
    assert.match(formatMatrixRoleDefsLoadFailureMessage("HTTP 403"), /No file downloaded/i);
    assert.match(formatMatrixRoleDefsEmptyMessage(), /No role definitions returned/i);
    assert.match(formatMatrixRoleDefsEmptyMessage(), /No file downloaded/i);
  });
});

describe("formatMatrixItemScanFailureMessage", () => {
  it("includes the item label and says no file was downloaded", () => {
    const msg = formatMatrixItemScanFailureMessage(
      "/sites/hr/Shared Documents/secret.docx",
      new Error("HTTP 403")
    );
    assert.match(msg, /secret\.docx/);
    assert.match(msg, /HTTP 403/);
    assert.match(msg, /No file downloaded/i);
  });
});

describe("canReportExportCsvMatrixSuccess", () => {
  it("blocks success when role names failed or an item scan error was recorded", () => {
    assert.equal(
      canReportExportCsvMatrixSuccess({ roleNamesOk: true, itemScanError: null }),
      true
    );
    assert.equal(
      canReportExportCsvMatrixSuccess({
        roleNamesOk: true,
        itemScanError: new Error("HTTP 403")
      }),
      false
    );
    assert.equal(
      canReportExportCsvMatrixSuccess({ roleNamesOk: false, itemScanError: null }),
      false
    );
  });
});

describe("exportCSV.js variants fail closed on matrix role/item scan errors", () => {
  for (const { rel, source } of exportSources) {
    it(`${rel} inlines integrity helpers and uses them`, () => {
      assert.match(source, /function formatMatrixRoleDefsLoadFailureMessage\s*\(/);
      assert.match(source, /function formatMatrixRoleDefsEmptyMessage\s*\(/);
      assert.match(source, /function formatMatrixItemScanFailureMessage\s*\(/);
      assert.match(source, /formatMatrixRoleDefsLoadFailureMessage\s*\(/);
      assert.match(source, /formatMatrixRoleDefsEmptyMessage\s*\(\s*\)/);
      assert.match(source, /formatMatrixItemScanFailureMessage\s*\(/);
      assert.match(source, /No file downloaded/);
    });

    it(`${rel} does not fall back to a hard-coded preferred role list`, () => {
      // Former silent truncation: empty roleNames → hard-coded preferred array.
      assert.doesNotMatch(
        source,
        /if\s*\(\s*roleNames(?:Ws)?\.length\s*===\s*0\s*\)\s*roleNames(?:Ws)?\s*=\s*\[\s*"Full Control"/
      );
    });

    it(`${rel} does not swallow unique-item RoleAssignments failures`, () => {
      // Former fail-open concurrent worker catch that continued and reportDone(true).
      assert.doesNotMatch(
        source,
        /\.catch\s*\(\s*function\s*\(\s*\)\s*\{\s*completedCount\+\+;\s*work\(\);\s*\}\s*\)/
      );
      assert.match(source, /itemScanError/);
      assert.match(
        source,
        /reportDone\s*\(\s*false\s*,\s*formatMatrixItemScanFailureMessage/
      );
    });
  }
});
