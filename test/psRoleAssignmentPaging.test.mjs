import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const matrixPs = readFileSync(
  join(root, "scripts/Export-SitePermissionsMatrix.ps1"),
  "utf8"
);

function functionBody(source, name) {
  const start = source.indexOf(`function ${name}`);
  assert.notEqual(start, -1, `missing function ${name}`);
  const next = source.indexOf("\nfunction ", start + 1);
  return next === -1 ? source.slice(start) : source.slice(start, next);
}

describe("Get-RoleAssignments pages past the first REST response", () => {
  const body = functionBody(matrixPs, "Get-RoleAssignments");

  it("requests $top=5000 instead of the SharePoint default of 100", () => {
    assert.match(body, /roleassignments\?[^'"\n]*\$top=5000/);
    assert.match(body, /items\(\$ItemId\)\/roleassignments\?[^'"\n]*\$top=5000/);
  });

  it("follows OData nextLink until the collection is exhausted", () => {
    assert.match(body, /while\s+\(\$url\)/);
    assert.match(body, /Get-ODataNextLink/);
  });

  it("uses -Raw so @odata.nextLink is not dropped by PnP's parsed object", () => {
    assert.match(body, /Invoke-SPRestGetWithRetry\s+-Url\s+\$url\s+-Raw/);
  });

  it("keeps principals already collected if a later page fails", () => {
    assert.match(body, /Role assignments paging stopped/);
    assert.match(body, /\$collected\.Count -lt 1/);
  });
});

describe("Invoke-SPRestGet -Raw preserves nometadata nextLink for ACL paging", () => {
  const body = functionBody(matrixPs, "Invoke-SPRestGet");

  it("optional -Raw parses ConvertFrom-Json after Invoke-PnPSPRestMethod -Raw", () => {
    assert.match(body, /\[switch\]\s+\$Raw/);
    assert.match(body, /Invoke-PnPSPRestMethod\s+-Url\s+\$u\s+-Method\s+Get\s+-Raw/);
    assert.match(body, /ConvertFrom-Json/);
  });

  it("leaves the default parsed path in place for other callers", () => {
    assert.match(
      body,
      /return Invoke-PnPSPRestMethod\s+-Url\s+\$u\s+-Method\s+Get\s*$/m
    );
  });
});
