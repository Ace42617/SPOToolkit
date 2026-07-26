import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import {
  canReportMatrixSuccess,
  formatMatrixListScanFailureMessage,
  orderRoleNames,
  resolveRoleNamesFromDefs
} from "../lib/matrixExportIntegrity.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const exportSource = readFileSync(join(root, "permissionsMatrixExport.js"), "utf8");

const PREFERRED = [
  "Full Control",
  "Design",
  "Edit",
  "Contribute",
  "Read",
  "View Only"
];

describe("resolveRoleNamesFromDefs", () => {
  it("orders preferred roles first and keeps custom roles", () => {
    const resolved = resolveRoleNamesFromDefs(
      [{ Name: "Records Manager" }, { Name: "Read" }, { Name: "Full Control" }],
      PREFERRED
    );
    assert.equal(resolved.ok, true);
    assert.deepEqual(resolved.roleNames, [
      "Full Control",
      "Read",
      "Records Manager"
    ]);
  });

  it("fails closed when roledefinitions returns no names (no preferred-list fallback)", () => {
    const resolved = resolveRoleNamesFromDefs([], PREFERRED);
    assert.equal(resolved.ok, false);
    assert.match(resolved.error, /No role definitions returned/i);
    assert.deepEqual(resolved.roleNames, []);
  });
});

describe("orderRoleNames", () => {
  it("dedupes and preserves non-preferred roles", () => {
    assert.deepEqual(orderRoleNames(["Edit", "Edit", "Custom"], ["Edit", "Read"]), [
      "Edit",
      "Custom"
    ]);
  });
});

describe("canReportMatrixSuccess / formatMatrixListScanFailureMessage", () => {
  it("blocks success when any list scan error was recorded", () => {
    assert.equal(canReportMatrixSuccess({ roleNamesOk: true, scanErrors: [] }), true);
    assert.equal(
      canReportMatrixSuccess({
        roleNamesOk: true,
        scanErrors: [{ listTitle: "Contracts" }]
      }),
      false
    );
    assert.equal(canReportMatrixSuccess({ roleNamesOk: false, scanErrors: [] }), false);
  });

  it("tells the operator that no workbook was downloaded", () => {
    const msg = formatMatrixListScanFailureMessage("HR Docs", new Error("HTTP 403"));
    assert.match(msg, /HR Docs/);
    assert.match(msg, /HTTP 403/);
    assert.match(msg, /No file downloaded/i);
  });
});

describe("permissionsMatrixExport.js stays fail-closed", () => {
  it("does not fall back to PREFERRED_ROLES when roledefinitions fail or return empty", () => {
    assert.match(exportSource, /function fetchRoleNames\s*\(/);
    assert.match(exportSource, /Failed to load role definitions/);
    assert.match(exportSource, /No role definitions returned for this site/);
    // The silent empty→PREFERRED fallback must stay gone.
    assert.doesNotMatch(
      exportSource,
      /if\s*\(\s*roleNames\.length\s*===\s*0\s*\)\s*roleNames\s*=\s*PREFERRED_ROLES\.slice\s*\(\s*\)/
    );
  });

  it("does not skip failed lists and still reportDone(true)", () => {
    assert.match(
      exportSource,
      /throw new Error\(\s*formatMatrixListScanFailureMessage\(\s*listTitle\s*,\s*listErr\s*\)\s*\)/
    );
    assert.match(exportSource, /Permissions matrix failed while scanning/);
    assert.match(exportSource, /No file downloaded/);
    // Former fail-open progress copy that continued the run.
    assert.doesNotMatch(
      exportSource,
      /Permissions matrix:\s*Skipped\s*"\s*\+\s*listTitle/
    );
  });
});
