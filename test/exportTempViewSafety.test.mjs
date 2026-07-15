import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const exportFiles = [
  new URL("../exportCSV.js", import.meta.url),
  new URL("../SP-Developer-Toolkit-Experimental/exportCSV.js", import.meta.url),
];

describe("temporary export view safety", () => {
  for (const file of exportFiles) {
    it(`does not replace an existing RPC view in ${file.pathname}`, async () => {
      const source = await readFile(file, "utf8");
      const setupStart = source.indexOf("async function getOrCreateRpcViewForDownload()");
      const setupEnd = source.indexOf("function mergeRequiredFields", setupStart);
      assert.notEqual(setupStart, -1);
      assert.notEqual(setupEnd, -1);

      const setupSource = source.slice(setupStart, setupEnd);
      assert.doesNotMatch(setupSource, /deleteViewById/);
      assert.doesNotMatch(setupSource, /\/views\?\$select=Id,Title/);
      assert.match(setupSource, /Title:\s*RPC_VIEW_TITLE/);
    });

    it(`uses a bounded export-owned view title in ${file.pathname}`, async () => {
      const source = await readFile(file, "utf8");
      const helperSource = source.match(
        /function buildTemporaryViewTitle\(baseTitle\) \{[\s\S]*?\n  \}/
      )?.[0];
      assert.ok(helperSource, "temporary title helper is present");

      const buildTitle = Function(`return (${helperSource});`)();
      const title = buildTitle("R".repeat(300));
      assert.equal(title.length, 255);
      assert.match(title, /-SPOToolkit-[a-z0-9]+-[a-z0-9]+$/);
      assert.notEqual(buildTitle("RPC"), "RPC");
    });

    it(`clears the paging filter when a working view is retained in ${file.pathname}`, async () => {
      const source = await readFile(file, "utf8");
      assert.match(
        source,
        /if \(DELETE_VIEW_AT_END\)[\s\S]*?deleteViewById\(rpcViewId\)[\s\S]*?else \{[\s\S]*?setViewQueryById\(rpcViewId, ""\)/
      );
    });
  }
});
