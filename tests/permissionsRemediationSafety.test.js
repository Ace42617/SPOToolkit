import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

const script = readFileSync(new URL("../scripts/Invoke-PermissionsMatrixRemediation.ps1", import.meta.url), "utf8");

const removeSharingLinkCommands = script
  .split(/\r?\n/)
  .map((line) => line.trim())
  .filter((line) => /^Remove-PnP(?:Folder|File)SharingLink\b/.test(line));

assert.ok(removeSharingLinkCommands.length > 0, "expected sharing-link removal commands");
for (const command of removeSharingLinkCommands) {
  assert.match(command, /\s-Identity\s+\$shareId\b/, `sharing-link removal must target a concrete ShareId: ${command}`);
}

const removeExternalUserFunctionMatch = script.match(
  /function Invoke-RemoveExternalUserTarget \{[\s\S]*?(?=^function Invoke-RemoveSharingLinkTarget \{)/m,
);
assert.ok(removeExternalUserFunctionMatch, "expected external-user remediation function");

const removeExternalUserFunction = removeExternalUserFunctionMatch[0];
const fileFolderBranchStart = removeExternalUserFunction.indexOf("'^(File|Folder|List item)$' {");
const defaultBranchStart = removeExternalUserFunction.indexOf("default {", fileFolderBranchStart);
assert.ok(fileFolderBranchStart >= 0 && defaultBranchStart > fileFolderBranchStart, "expected file/folder/list item remediation branch");

const fileFolderBranch = removeExternalUserFunction.slice(fileFolderBranchStart, defaultBranchStart);
assert.match(fileFolderBranch, /ScopeKind ListItem\b/, "item remediation should remove the selected item assignment");
assert.doesNotMatch(fileFolderBranch, /ScopeKind List\b/, "item remediation must not cascade to parent list/library");
assert.doesNotMatch(fileFolderBranch, /ScopeKind Web\b/, "item remediation must not cascade to parent site/web");

assert.doesNotMatch(
  script,
  /Remove (?:site|library)-level grants for external user/,
  "plan builder must not synthesize hidden parent-scope remediation targets",
);
