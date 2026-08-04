/**
 * Regression: borrowed OWSSVR library views must never be mutated via
 * RemoveAllViewFields / setViewFieldsBatch while reported as read-only.
 */
import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { shouldConfigureOwssvrExportView } from "../lib/owssvrBorrowedView.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const rootExport = readFileSync(join(root, "exportCSV.js"), "utf8");

describe("shouldConfigureOwssvrExportView", () => {
  it("allows mutation only for owned, non-read-only export views", () => {
    assert.equal(
      shouldConfigureOwssvrExportView({ ownView: true, readOnlyView: false }),
      true
    );
  });

  it("blocks mutation for borrowed / read-only library views", () => {
    assert.equal(
      shouldConfigureOwssvrExportView({ ownView: false, readOnlyView: true }),
      false
    );
    assert.equal(shouldConfigureOwssvrExportView({ readOnlyView: true }), false);
    assert.equal(shouldConfigureOwssvrExportView({ ownView: false }), false);
  });

  it("fails closed on missing view info", () => {
    assert.equal(shouldConfigureOwssvrExportView(null), false);
    assert.equal(shouldConfigureOwssvrExportView(undefined), false);
  });
});

describe("exportCSV.js borrowed view stays read-only (source sync)", () => {
  it("does not call tryConfigureRpcView after borrowing a library view", () => {
    assert.match(rootExport, /async function getOrCreateRpcViewForDownload\s*\(/);
    assert.match(rootExport, /resolveBorrowedOwssvrViewId\s*\(/);
    assert.match(rootExport, /readOnlyView:\s*true/);

    const borrowedBlock = rootExport.match(
      /var borrowed = await resolveBorrowedOwssvrViewId\(\);[\s\S]*?readOnlyView:\s*true/
    );
    assert.ok(
      borrowedBlock,
      "expected borrowed-view return block in getOrCreateRpcViewForDownload"
    );
    // Strip comments so explanatory text cannot false-positive the guard.
    const codeOnly = borrowedBlock[0].replace(/\/\*[\s\S]*?\*\//g, "").replace(/\/\/.*$/gm, "");
    assert.doesNotMatch(codeOnly, /tryConfigureRpcView\s*\(/);
    assert.doesNotMatch(codeOnly, /setViewFieldsBatch\s*\(/);
    assert.doesNotMatch(codeOnly, /RemoveAllViewFields/);
  });

  it("still configures owned RPC views before export", () => {
    assert.match(
      rootExport,
      /await tryConfigureRpcView\(created\.viewId,\s*fieldNames\)/
    );
    assert.match(
      rootExport,
      /await tryConfigureRpcView\(existingRpc\.id,\s*fieldNames\)/
    );
  });
});
