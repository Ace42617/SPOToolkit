import { describe, it } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const remediationScript = fs.readFileSync(
  path.join(root, "scripts", "Invoke-PermissionsMatrixRemediation.ps1"),
  "utf8",
);
const exportScript = fs.readFileSync(
  path.join(root, "scripts", "Export-SitePermissionsMatrix.ps1"),
  "utf8",
);

function getFunction(source, name) {
  const marker = `function ${name}`;
  const start = source.indexOf(marker);
  assert.notEqual(start, -1, `${name} should exist`);
  const next = source.indexOf("\nfunction ", start + marker.length);
  return source.slice(start, next === -1 ? source.length : next);
}

describe("permissions remediation safety invariants", () => {
  it("refuses tenant admin URLs before mutating permissions", () => {
    const connect = getFunction(remediationScript, "Connect-PnPSiteIfNeeded");

    assert.match(connect, /Test-IsSharePointTenantAdminUrl\s+\$target/);
    assert.match(connect, /Refusing to run remediation against tenant admin URL/);
  });

  it("does not implicitly cascade item external-user removal to parent scopes", () => {
    const removeTarget = getFunction(remediationScript, "Invoke-RemoveExternalUserTarget");
    const fileCase = removeTarget.slice(removeTarget.indexOf("'^(File|Folder|List item)$'"));

    assert.match(fileCase, /ScopeKind ListItem/);
    assert.doesNotMatch(fileCase, /ScopeKind List\s+-List \$list/);
    assert.doesNotMatch(fileCase, /ScopeKind Web/);
  });

  it("exports and consumes explicit matrix row site URLs", () => {
    const newMatrixRow = getFunction(exportScript, "New-MatrixRow");
    const exportWorkbook = getFunction(exportScript, "Export-Workbook");
    const getSiteUrl = getFunction(remediationScript, "Get-SiteUrlForRow");

    assert.match(newMatrixRow, /'Site URL'\s+=\s+\$S\.ConnectedWebUrl/);
    assert.match(exportWorkbook, /'Site Name', 'Site URL', 'Name'/);
    assert.match(getSiteUrl, /\$Row\.'Site URL'/);
    assert.match(remediationScript, /'Site Name', 'Site URL', 'Item path'/);
  });
});
