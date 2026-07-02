import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const remediationScript = new URL("../scripts/Invoke-PermissionsMatrixRemediation.ps1", import.meta.url);
const backgroundScript = new URL("../background.js", import.meta.url);
const matrixWorkerLockScript = new URL("../matrixWorkerLock.js", import.meta.url);

function extractFunction(source, name) {
  const start = source.indexOf(`function ${name}`);
  assert.notEqual(start, -1, `${name} should exist`);
  const next = source.indexOf("\nfunction ", start + 1);
  return next === -1 ? source.slice(start) : source.slice(start, next);
}

test("sharing-link remediation refuses rows without a concrete ShareId", async () => {
  const source = await readFile(remediationScript, "utf8");
  const planBuilder = extractFunction(source, "Build-SharingLinkPlan");
  const executor = extractFunction(source, "Invoke-RemoveSharingLinkTarget");

  assert.match(planBuilder, /\$shareId = \(\[string\]\$row\.ShareId\)\.Trim\(\)/);
  assert.match(planBuilder, /if \(-not \$shareId -or -not \$rel\) \{ continue \}/);
  assert.match(executor, /if \(-not \$shareId\) \{\s*throw "Refusing to remove sharing links/s);

  const removeCalls = executor
    .split("\n")
    .filter((line) => /Remove-PnP(?:Folder|File)SharingLink/.test(line));
  assert.ok(removeCalls.length >= 2, "expected file and folder sharing-link removal calls");
  for (const call of removeCalls) {
    assert.match(call, / -Identity \$shareId\b/, `unsafe remove call: ${call}`);
  }
});

test("second-stage recycle-bin handoff keeps the service worker alive until shutdown", async () => {
  const source = await readFile(backgroundScript, "utf8");
  const opener = extractFunction(source, "openSecondStageRecycleBinSequence");

  assert.match(opener, /function openSecondStageRecycleBinSequence\(tabId, firstStageUrl, secondStageUrl, keepalivePort\)/);
  assert.match(opener, /let portToClose = keepalivePort/);
  assert.match(opener, /function disconnectKeepalive\(\)/);
  assert.match(opener, /disconnectKeepalive\(\);/);
});

test("permissions matrix completion dismisses the worker overlay without closing the tab", async () => {
  const background = await readFile(backgroundScript, "utf8");
  const workerLock = await readFile(matrixWorkerLockScript, "utf8");

  assert.doesNotMatch(background, /chrome\.tabs\.remove\(workerTabId/);
  assert.match(workerLock, /returns to normal when finished/);
  assert.match(workerLock, /setTimeout\(hide, dismissMs\)/);
  assert.doesNotMatch(workerLock, /Closing this tab in/);
});
