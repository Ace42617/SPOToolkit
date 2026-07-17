import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const formatterPaths = [
  new URL("../view-formatter.js", import.meta.url),
  new URL("../SP-Developer-Toolkit-Experimental/view-formatter.js", import.meta.url),
];

for (const formatterPath of formatterPaths) {
  const label = formatterPath.pathname.includes("Experimental") ? "Experimental" : "root";

  test(`${label} View Formatter ignores stale view loads`, async () => {
    const source = await readFile(formatterPath, "utf8");

    assert.match(source, /const loadGeneration = \+\+viewLoadGeneration;/);
    assert.match(
      source,
      /if \(loadGeneration !== viewLoadGeneration\) return false;[\s\S]*if \(!res \|\| !res\.ok\)/
    );
  });

  test(`${label} View Formatter saves only the fully loaded selected view`, async () => {
    const source = await readFile(formatterPath, "utf8");

    assert.match(
      source,
      /const targetViewId = normGuid\(selectedViewId\);[\s\S]*!sameGuid\(viewSelectEl\.value, targetViewId\)/
    );
    assert.match(
      source,
      /const canSave =[\s\S]*selectedOptionMatches[\s\S]*!viewLoadPending[\s\S]*!viewSavePending;/
    );
    assert.match(
      source,
      /const ctx = await sendToSpTab\(\{ action: "getViewFormatContext" \}\);[\s\S]*const viewId = targetViewId;/
    );
    assert.match(
      source,
      /!sameGuid\(liveListId, contextListId\)[\s\S]*normalizedSitePath\(liveSitePath\) !== normalizedSitePath\(contextSitePath\)/
    );
    assert.match(source, /const listId = contextListId;[\s\S]*const sitePath = contextSitePath;/);
    assert.doesNotMatch(source, /selectedViewId \|\| ctx\.viewId/);
  });
}
