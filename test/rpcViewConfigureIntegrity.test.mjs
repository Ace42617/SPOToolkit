/**
 * Regression: OWSSVR export must not report success after RPC view field setup
 * fails (RemoveAllViewFields succeeded, addViewField mostly failed).
 */
import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import {
  interpretTryConfigureRpcView,
  ownedRpcViewAfterConfigure
} from "../lib/rpcViewConfigureIntegrity.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const rootExport = readFileSync(join(root, "exportCSV.js"), "utf8");

describe("interpretTryConfigureRpcView", () => {
  it("passes through a successful setViewFieldsBatch", () => {
    assert.deepEqual(interpretTryConfigureRpcView({ ok: true }), {
      ok: true,
      fieldsConfigured: true
    });
  });

  it("fails closed when setViewFieldsBatch returns ok:false", () => {
    const r = interpretTryConfigureRpcView({
      ok: false,
      error: "View field setup failed (403 or permission denied). Try refreshing the page and run the export again."
    });
    assert.equal(r.ok, false);
    assert.equal(r.fieldsConfigured, false);
    assert.match(r.error, /View field setup failed/);
  });

  it("fails closed when setViewFieldsBatch result is missing", () => {
    const r = interpretTryConfigureRpcView(null);
    assert.equal(r.ok, false);
    assert.equal(r.fieldsConfigured, false);
    assert.match(r.error, /No file downloaded/i);
  });
});

describe("ownedRpcViewAfterConfigure", () => {
  const payload = {
    ok: true,
    viewId: "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee",
    ownView: true,
    readOnlyView: false,
    fieldNames: ["ID", "Title", "FileLeafRef"]
  };

  it("returns the owned-view payload when configure succeeded", () => {
    assert.deepEqual(
      ownedRpcViewAfterConfigure({ ok: true, fieldsConfigured: true }, payload),
      payload
    );
  });

  it("does not proceed to OWSSVR export when configure failed", () => {
    // Concrete trigger: document-library export, RemoveAllViewFields succeeds,
    // then a majority of addViewField calls fail (403/throttle). Former code
    // returned { ok: true } from tryConfigureRpcView and ignored it.
    const failed = ownedRpcViewAfterConfigure(
      { ok: false, fieldsConfigured: false, error: "View field setup failed (403 or permission denied)." },
      payload
    );
    assert.equal(failed.ok, false);
    assert.match(failed.error, /View field setup failed/);
    assert.equal(failed.viewId, undefined);
  });
});

describe("exportCSV.js RPC configure stays fail-closed (source sync)", () => {
  it("tryConfigureRpcView returns ok:false when setViewFieldsBatch fails", () => {
    assert.match(rootExport, /async function tryConfigureRpcView\s*\(/);
    assert.match(
      rootExport,
      /if\s*\(\s*sf\.ok\s*\)\s*return\s*\{\s*ok:\s*true,\s*fieldsConfigured:\s*true\s*\}/
    );
    assert.match(
      rootExport,
      /return\s*\{\s*ok:\s*false,\s*fieldsConfigured:\s*false,\s*error:\s*sf\.error/
    );
    assert.match(rootExport, /No file downloaded/);
    assert.doesNotMatch(
      rootExport,
      /return\s*\{\s*ok:\s*true,\s*fieldsConfigured:\s*false/
    );
  });

  it("owned RPC paths fail closed when tryConfigureRpcView fails", () => {
    assert.match(rootExport, /async function getOrCreateRpcViewForDownload\s*\(/);
    assert.match(
      rootExport,
      /var cfgExisting = await tryConfigureRpcView\(existingRpc\.id,\s*fieldNames\);\s*if\s*\(\s*!cfgExisting\.ok\s*\)\s*return cfgExisting;/
    );
    assert.match(
      rootExport,
      /var cfgCreated = await tryConfigureRpcView\(created\.viewId,\s*fieldNames\);\s*if\s*\(\s*!cfgCreated\.ok\s*\)\s*return cfgCreated;/
    );
    assert.match(
      rootExport,
      /var cfgPersonal = await tryConfigureRpcView\(created\.viewId,\s*fieldNames\);\s*if\s*\(\s*!cfgPersonal\.ok\s*\)\s*return cfgPersonal;/
    );
  });

  it("does not mutate a borrowed library view after configure is fail-closed", () => {
    const borrowedBlock = rootExport.match(
      /var borrowed = await resolveBorrowedOwssvrViewId\(\);[\s\S]*?readOnlyView:\s*true/
    );
    assert.ok(
      borrowedBlock,
      "expected borrowed-view return block in getOrCreateRpcViewForDownload"
    );
    const codeOnly = borrowedBlock[0]
      .replace(/\/\*[\s\S]*?\*\//g, "")
      .replace(/\/\/.*$/gm, "");
    assert.doesNotMatch(codeOnly, /tryConfigureRpcView\s*\(/);
    assert.doesNotMatch(codeOnly, /setViewFieldsBatch\s*\(/);
  });
});
