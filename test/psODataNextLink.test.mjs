import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { getODataNextLinkFromRestPayload } from "../lib/psODataNextLink.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const sharingPs = readFileSync(
  join(root, "scripts/Export-SiteSharingLinks.ps1"),
  "utf8"
);
const matrixPs = readFileSync(
  join(root, "scripts/Export-SitePermissionsMatrix.ps1"),
  "utf8"
);

describe("getODataNextLinkFromRestPayload", () => {
  it("reads nometadata @odata.nextLink (SharePoint default JSON light)", () => {
    assert.equal(
      getODataNextLinkFromRestPayload({
        value: [{ Id: 1 }],
        "@odata.nextLink": "https://contoso.sharepoint.com/_api/web/lists/items?$skiptoken=Paged=TRUE"
      }),
      "https://contoso.sharepoint.com/_api/web/lists/items?$skiptoken=Paged=TRUE"
    );
  });

  it("reads odata.nextLink without @ (minimalmetadata / PnP NoteProperty)", () => {
    assert.equal(
      getODataNextLinkFromRestPayload({
        value: [{ Id: 1 }],
        "odata.nextLink": "https://contoso.sharepoint.com/_api/web/lists/items?$skip=2000"
      }),
      "https://contoso.sharepoint.com/_api/web/lists/items?$skip=2000"
    );
  });

  it("reads verbose d.__next", () => {
    assert.equal(
      getODataNextLinkFromRestPayload({
        d: { results: [{ Id: 1 }], __next: "https://contoso.sharepoint.com/_api/web/lists/items?$skip=2000" }
      }),
      "https://contoso.sharepoint.com/_api/web/lists/items?$skip=2000"
    );
  });

  it("returns null when the page is complete", () => {
    assert.equal(getODataNextLinkFromRestPayload({ value: [{ Id: 1 }] }), null);
    assert.equal(getODataNextLinkFromRestPayload(null), null);
  });

  it("prefers @odata.nextLink when both names are present", () => {
    assert.equal(
      getODataNextLinkFromRestPayload({
        "@odata.nextLink": "https://a.example/_api/next-at",
        "odata.nextLink": "https://a.example/_api/next-plain"
      }),
      "https://a.example/_api/next-at"
    );
  });
});

describe("Export-SiteSharingLinks.ps1 pages past the first $top=2000 response", () => {
  it("parses REST with -Raw so @odata.nextLink is not dropped", () => {
    assert.match(sharingPs, /Invoke-PnPSPRestMethod\s+-Url\s+\$url\s+-Method\s+Get\s+-Raw/);
    assert.match(sharingPs, /ConvertFrom-Json/);
  });

  it("follows @odata.nextLink and odata.nextLink, not only odata.nextLink", () => {
    assert.match(sharingPs, /'@odata\.nextLink'/);
    assert.match(sharingPs, /'odata\.nextLink'/);
    assert.doesNotMatch(
      sharingPs,
      /\$url\s*=\s*\$response\.'odata\.nextLink'/
    );
  });
});

describe("Export-SitePermissionsMatrix.ps1 preserves OData nextLink on REST GET", () => {
  it("uses -Raw in Invoke-SPRestGet so Get-ODataNextLink can see @odata.nextLink", () => {
    assert.match(matrixPs, /function Invoke-SPRestGet\(/);
    const start = matrixPs.indexOf("function Invoke-SPRestGet(");
    const restGet = matrixPs.slice(start, start + 1200);
    assert.match(restGet, /-Raw/);
    assert.match(restGet, /ConvertFrom-Json/);
  });

  it("reads both nextLink property names", () => {
    assert.match(
      matrixPs,
      /function Get-ODataNextLink[\s\S]*?'@odata\.nextLink'[\s\S]*?'odata\.nextLink'/
    );
  });
});
