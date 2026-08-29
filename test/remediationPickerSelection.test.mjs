import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

const source = readFileSync(
  new URL("../scripts/Invoke-PermissionsMatrixRemediation.ps1", import.meta.url),
  "utf8",
);

function between(start, end) {
  const from = source.indexOf(start);
  const to = source.indexOf(end, from + start.length);
  assert.notEqual(from, -1, `Missing ${start}`);
  assert.notEqual(to, -1, `Missing boundary after ${start}`);
  return source.slice(from, to);
}

test("remediation exclusions survive list sorting and filtering", () => {
  const saveState = between(
    "function Save-TargetCheckState",
    "function Save-UniqueTreeCheckState",
  );
  assert.doesNotMatch(saveState, /\$unchecked\.Clear\(\)/);
  assert.match(saveState, /Get-TrackerChecked \$t[\s\S]*\$unchecked\.Remove/);
  assert.match(saveState, /else \{[\s\S]*\$unchecked\.Add/);

  const saveUniqueState = between(
    "function Save-UniqueTreeCheckState",
    "function Get-SelectedPlanItems",
  );
  assert.doesNotMatch(
    saveUniqueState,
    /\$(?:parentUnchecked|childUnchecked)\.Clear\(\)/,
  );
  assert.match(
    saveUniqueState,
    /\$unchecked = if \(\$t\.IsParent\)[\s\S]*\$unchecked\.Remove/,
  );
  assert.match(saveUniqueState, /else \{[\s\S]*\$unchecked\.Add/);

  const render = between(
    "function Render-PlanListView",
    "function Render-ExternalListView",
  );
  assert.match(
    render,
    /\$unchecked = \$script:RemediationPickerUncheckedKeys\[\$TabKey\]/,
  );
  assert.match(
    render,
    /-Checked \(-not \$unchecked\.Contains\(\[string\]\$row\.PlanKey\)\)/,
  );
  assert.doesNotMatch(render, /-Checked \$row\.Checked/);

  const allExternal = between(
    "$btnAllExt.Add_Click",
    "$btnNoneExt.Add_Click",
  );
  const allLinks = between(
    "$btnAllLinks.Add_Click",
    "$btnNoneLinks.Add_Click",
  );
  const allUnique = between(
    "function Set-UniqueTreeChecks",
    "$btnNoneUnique.Add_Click",
  );
  for (const handler of [allExternal, allLinks, allUnique]) {
    assert.doesNotMatch(handler, /RemediationPickerUncheckedKeys[\s\S]*\.Clear\(\)/);
  }
  assert.match(allUnique, /Save-UniqueTreeCheckState/);
});
