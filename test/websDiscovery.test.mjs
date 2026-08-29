import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, it } from "node:test";
import {
  CHILD_WEBS_PAGE_SIZE,
  childWebsApiUrl,
  extractODataNextLink,
  extractODataValue,
  fetchAllChildWebs
} from "../lib/websDiscovery.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");

describe("websDiscovery helpers", () => {
  it("builds a paged child-webs URL", () => {
    assert.equal(
      childWebsApiUrl("https://contoso.sharepoint.com/sites/portal/"),
      "https://contoso.sharepoint.com/sites/portal/_api/web/webs?$select=Title,ServerRelativeUrl,Url&$top=" +
        CHILD_WEBS_PAGE_SIZE
    );
    assert.equal(CHILD_WEBS_PAGE_SIZE, 500);
  });

  it("extracts OData values and next links", () => {
    assert.deepEqual(extractODataValue({ value: [{ a: 1 }] }), [{ a: 1 }]);
    assert.deepEqual(extractODataValue({ results: [{ b: 2 }] }), [{ b: 2 }]);
    assert.deepEqual(extractODataValue(null), []);
    assert.equal(
      extractODataNextLink({ "@odata.nextLink": "https://x/next" }),
      "https://x/next"
    );
    assert.equal(
      extractODataNextLink({ "odata.nextLink": "https://x/legacy" }),
      "https://x/legacy"
    );
    assert.equal(extractODataNextLink({ value: [] }), null);
  });

  it("follows @odata.nextLink across child-web pages", async () => {
    const calls = [];
    const pages = {
      "https://contoso.sharepoint.com/sites/hub/_api/web/webs?$select=Title,ServerRelativeUrl,Url&$top=500": {
        value: Array.from({ length: 500 }, (_, i) => ({
          ServerRelativeUrl: "/sites/hub/w" + i
        })),
        "@odata.nextLink": "https://contoso.sharepoint.com/sites/hub/_api/web/webs?$skiptoken=paged&$top=500"
      },
      "https://contoso.sharepoint.com/sites/hub/_api/web/webs?$skiptoken=paged&$top=500": {
        value: Array.from({ length: 50 }, (_, i) => ({
          ServerRelativeUrl: "/sites/hub/extra" + i
        }))
      }
    };
    async function fetchJson(url) {
      calls.push(url);
      const j = pages[url];
      if (!j) throw new Error("unexpected url " + url);
      return j;
    }
    const items = await fetchAllChildWebs(
      fetchJson,
      "https://contoso.sharepoint.com/sites/hub",
      "application/json;odata=nometadata"
    );
    assert.equal(items.length, 550);
    assert.equal(calls.length, 2);
    assert.equal(items[0].ServerRelativeUrl, "/sites/hub/w0");
    assert.equal(items[549].ServerRelativeUrl, "/sites/hub/extra49");
  });

  it("applies normalizeUrl to nextLink before the next fetch", async () => {
    const calls = [];
    async function fetchJson(url) {
      calls.push(url);
      if (calls.length === 1) {
        return {
          value: [{ ServerRelativeUrl: "/sites/a/one" }],
          "@odata.nextLink": "https://other-host/_api/web/webs?$skiptoken=x"
        };
      }
      return { value: [{ ServerRelativeUrl: "/sites/a/two" }] };
    }
    const items = await fetchAllChildWebs(
      fetchJson,
      "https://contoso.sharepoint.com/sites/a",
      "application/json;odata=nometadata",
      (url) => "https://contoso.sharepoint.com" + new URL(url).pathname + new URL(url).search
    );
    assert.equal(items.length, 2);
    assert.equal(
      calls[1],
      "https://contoso.sharepoint.com/_api/web/webs?$skiptoken=x"
    );
  });
});

describe("websDiscovery source sync", () => {
  const sources = [
    ["permissionsMatrixExport.js", readFileSync(join(root, "permissionsMatrixExport.js"), "utf8")],
    ["matrixScanPlan.js", readFileSync(join(root, "matrixScanPlan.js"), "utf8")],
    ["reportListPlan.js", readFileSync(join(root, "reportListPlan.js"), "utf8")]
  ];

  for (const [name, source] of sources) {
    it(`${name} pages /_api/web/webs via nextLink instead of a single $top=500 page`, () => {
      assert.match(source, /async function fetchAllChildWebs\(/);
      assert.match(
        source,
        /\/_api\/web\/webs\?\$select=Title,ServerRelativeUrl,Url&\$top=500/
      );
      assert.match(source, /\["@odata\.nextLink"\]/);
      assert.match(source, /while\s*\(\s*next\s*\)/);
      // Must not keep the old one-shot pattern that drops sibling webs beyond page 1.
      assert.doesNotMatch(
        source,
        /var subs = await fetchJson\([^;]+\/_api\/web\/webs\?\$select=Title,ServerRelativeUrl,Url&\$top=500[^;]*;\s*\n\s*var (items|subItems) = subs\.value/
      );
    });
  }
});
