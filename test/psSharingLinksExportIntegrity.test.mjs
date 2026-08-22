import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import {
  canExportPsSharingLinksCsv,
  formatPsSharingLinksListScanFailureMessage
} from "../lib/psSharingLinksExportIntegrity.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const psSource = readFileSync(
  join(root, "scripts/Export-SiteSharingLinks.ps1"),
  "utf8"
);

describe("formatPsSharingLinksListScanFailureMessage", () => {
  it("names the list, detail, and that no CSV was downloaded", () => {
    const msg = formatPsSharingLinksListScanFailureMessage(
      "Contracts",
      new Error("HTTP 403")
    );
    assert.match(msg, /Contracts/);
    assert.match(msg, /HTTP 403/);
    assert.match(msg, /No CSV downloaded/i);
  });
});

describe("canExportPsSharingLinksCsv", () => {
  it("allows export only after a fully successful scan", () => {
    assert.equal(canExportPsSharingLinksCsv({ scanCompletedOk: true }), true);
    assert.equal(canExportPsSharingLinksCsv({ scanCompletedOk: false }), false);
    assert.equal(
      canExportPsSharingLinksCsv({
        scanCompletedOk: true,
        itemScanFailures: [{ listTitle: "Contracts" }]
      }),
      false
    );
  });
});

describe("Export-SiteSharingLinks.ps1 stays fail-closed on list REST scan", () => {
  it("rethrows list REST failures instead of continuing to the next list", () => {
    assert.match(
      psSource,
      /Sharing-links export failed while scanning list/
    );
    assert.match(psSource, /No CSV downloaded/);
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
      /REST enumeration failed for list[\s\S]{0,160}break/
    );
  });

  it("keeps CSV write and Done behind the scan-completed guard", () => {
    const errGuardIdx = psSource.indexOf("if ($null -ne $scanError)");
    const okGuardIdx = psSource.indexOf("if (-not $scanCompletedOk)");
    const exportIdx = psSource.indexOf("Export-Csv");
    const doneIdx = psSource.indexOf('Write-Host "Done."');
    assert.ok(errGuardIdx > 0, "expected scanError rethrow guard");
    assert.ok(okGuardIdx > errGuardIdx, "expected scanCompletedOk guard");
    assert.ok(exportIdx > okGuardIdx, "Export-Csv must follow the guards");
    assert.ok(doneIdx > exportIdx, "success log must follow Export-Csv");
  });
});
