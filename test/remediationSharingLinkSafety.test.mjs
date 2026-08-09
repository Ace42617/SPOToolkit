import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const exportSource = await readFile(new URL("../permissionsMatrixExport.js", import.meta.url), "utf8");
const remediationSource = await readFile(
  new URL("../scripts/Invoke-PermissionsMatrixRemediation.ps1", import.meta.url),
  "utf8"
);

describe("sharing-link remediation safety", () => {
  it("exports ShareId values from SharePoint linkDetails payloads", () => {
    const shareIdAssignment = exportSource.match(/var shareId = (?<expr>[^;]+);/);
    assert.ok(shareIdAssignment, "pushSharingLinkRow should derive a ShareId before writing the row");

    const expr = shareIdAssignment.groups.expr;
    assert.ok(expr.includes("details.shareId"), "lowercase linkDetails.shareId should be exported");
    assert.ok(expr.includes("details.ShareId"), "PascalCase linkDetails.ShareId should be exported");
    assert.ok(
      expr.indexOf("details.ShareId") < expr.indexOf("link.id"),
      "nested ShareId should be preferred before generic top-level link ids"
    );
  });

  it("does not plan sharing-link removals when ShareId is blank", () => {
    assert.match(
      remediationSource,
      /\$shareId = \(\[string\]\$row\.ShareId\)\.Trim\(\)/,
      "planner should trim the ShareId value from the report row"
    );
    assert.match(
      remediationSource,
      /if \(-not \$shareId\) \{[\s\S]*?Skipping sharing link[\s\S]*?continue[\s\S]*?\}/,
      "planner should skip rows that do not identify the exact sharing link"
    );
  });

  it("always passes -Identity to destructive PnP sharing-link removal cmdlets", () => {
    const removalLines = remediationSource
      .split(/\r?\n/)
      .filter((line) => /Remove-PnP(?:Folder|File)SharingLink/.test(line));

    assert.deepEqual(removalLines, [
      "        Remove-PnPFolderSharingLink -Folder $rel -Identity $shareId -Force -ErrorAction Stop",
      "        Remove-PnPFileSharingLink -FileUrl $rel -Identity $shareId -Force -ErrorAction Stop",
    ]);
  });

  it("fails closed at execution time if a malformed target has no ShareId", () => {
    assert.match(
      remediationSource,
      /if \(-not \$shareId\) \{[\s\S]*?Sharing link removal requires a ShareId[\s\S]*?throw[\s\S]*?\}/,
      "executor should reject malformed sharing-link targets before calling PnP"
    );
  });
});
