import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import {
  canExportPsMatrixWorkbook,
  formatPsMatrixListScanFailureMessage
} from "../lib/psMatrixExportIntegrity.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const psSource = readFileSync(
  join(root, "scripts/Export-SitePermissionsMatrix.ps1"),
  "utf8"
);

describe("formatPsMatrixListScanFailureMessage", () => {
  it("names the list, URL, detail, and that no workbook was downloaded", () => {
    const msg = formatPsMatrixListScanFailureMessage(
      "Contracts",
      "/sites/hr/Contracts",
      new Error("HTTP 403")
    );
    assert.match(msg, /Contracts/);
    assert.match(msg, /\/sites\/hr\/Contracts/);
    assert.match(msg, /HTTP 403/);
    assert.match(msg, /No workbook downloaded/i);
  });
});

describe("canExportPsMatrixWorkbook", () => {
  it("allows export only after a fully successful scan", () => {
    assert.equal(canExportPsMatrixWorkbook({ scanCompletedOk: true }), true);
    assert.equal(canExportPsMatrixWorkbook({ scanCompletedOk: false }), false);
    assert.equal(
      canExportPsMatrixWorkbook({
        scanCompletedOk: true,
        itemScanFailures: [{ listTitle: "Contracts" }]
      }),
      false
    );
  });
});

describe("Export-SitePermissionsMatrix.ps1 stays fail-closed on item scan", () => {
  it("rethrows list item-scan failures instead of continuing to the next list", () => {
    assert.match(
      psSource,
      /Permissions matrix failed while scanning list/
    );
    assert.match(psSource, /No workbook downloaded/);
    assert.match(psSource, /\$scanCompletedOk\s*=\s*\$true/);
    assert.match(psSource, /\$scanError\s*=\s*\$_/);
    assert.match(psSource, /if\s*\(\s*\$null\s*-ne\s*\$scanError\s*\)/);
    assert.match(
      psSource,
      /if\s*\(\s*-not\s*\$scanCompletedOk\s*\)/
    );
    // Former fail-open continue-after-warning must stay gone.
    assert.doesNotMatch(
      psSource,
      /Item scan failed for list[\s\S]{0,120}Continuing with next list/
    );
  });

  it("keeps Export-Workbook behind the scan-completed guard", () => {
    const errGuardIdx = psSource.indexOf("if ($null -ne $scanError)");
    const okGuardIdx = psSource.indexOf("if (-not $scanCompletedOk)");
    const exportIdx = psSource.indexOf("Export-Workbook -MatrixRows");
    const doneIdx = psSource.indexOf(
      "Done — permissions matrix export complete."
    );
    assert.ok(errGuardIdx > 0, "expected scanError rethrow guard");
    assert.ok(okGuardIdx > errGuardIdx, "expected scanCompletedOk guard");
    assert.ok(exportIdx > okGuardIdx, "Export-Workbook must follow the guards");
    assert.ok(doneIdx > exportIdx, "success log must follow Export-Workbook");
  });
});

