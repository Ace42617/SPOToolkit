import assert from "node:assert/strict";
import fs from "node:fs";
import test from "node:test";

function read(path) {
  return fs.readFileSync(new URL(path, import.meta.url), "utf8");
}

function extractFunction(source, signature) {
  const start = source.indexOf(signature);
  assert.notEqual(start, -1, `Missing ${signature}`);
  const bodyStart = source.indexOf("{", start);
  let depth = 0;
  for (let i = bodyStart; i < source.length; i += 1) {
    if (source[i] === "{") depth += 1;
    if (source[i] === "}") {
      depth -= 1;
      if (depth === 0) return source.slice(start, i + 1);
    }
  }
  throw new Error(`Unterminated ${signature}`);
}

test("an explicitly empty matrix site selection performs no discovery", async () => {
  const source = read("../permissionsMatrixExport.js");
  const functionSource = extractFunction(source, "async function fetchAllWebs");
  const fetchAllWebs = Function(
    "SELECTED_SITE_PATHS",
    `"use strict"; return (${functionSource});`
  )([]);

  assert.deepEqual(await fetchAllWebs("https://tenant.sharepoint.com/sites/example", true), []);
  assert.match(
    source,
    /SELECTED_SITE_PATHS\s*&&\s*SELECTED_SITE_PATHS\.length\s*===\s*0[\s\S]*?No sites selected/
  );
});

test("content injection preserves an explicitly loaded empty selection", () => {
  const source = read("../content.js");
  assert.match(
    source,
    /Array\.isArray\(message\.matrixSelectedPaths\)[\s\S]*?message\.matrixSelectionLoaded\s*===\s*true[\s\S]*?params\.matrixSelectedPaths\s*=\s*message\.matrixSelectedPaths/
  );
});

test("popup and Compass reject None before launching a matrix worker", () => {
  for (const path of ["../popup.js", "../compassToolkitPanels.js"]) {
    const source = read(path);
    assert.match(source, /matrixSelectionLoaded[\s\S]*?matrixSelectedPaths\.length\s*===\s*0/);
    assert.match(source, /Select at least one site before running the Permissions Matrix export/);
  }
});
