import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const text = (path) => readFile(new URL("../" + path, import.meta.url), "utf8");

test("View formatter frame header rules are not enabled statically", async () => {
  const manifest = JSON.parse(await text("manifest.json"));
  const resources = manifest.declarative_net_request?.rule_resources || [];
  assert.equal(
    resources.some((resource) => String(resource.path || "").includes("view-formatter-iframe")),
    false
  );
});

test("View formatter frame header rules are scoped to formatter tabs", async () => {
  const background = await text("background.js");
  assert.match(background, /SPOToolkitEnableViewFormatterFrameRules/);
  assert.match(background, /updateSessionRules/);
  assert.match(background, /tabIds/);
});

test("View formatter save fails closed for ambiguous or navigated context", async () => {
  const formatter = await text("view-formatter.js");
  assert.doesNotMatch(formatter, /DefaultView\?\$select=Id/);
  assert.match(formatter, /Could not determine the current view id/);
  assert.match(formatter, /has navigated since View formatter opened/);
});

test("Recycle-bin wait overlays use timestamp expiry", async () => {
  for (const file of ["recycleBinWaitBanner.js", "SP-Developer-Toolkit-Lite/recycleBinWaitBanner.js"]) {
    const banner = await text(file);
    assert.match(banner, /__SPO_TOOLKIT_RB_WAIT_TS__/);
    assert.match(banner, /maxAgeMs/);
    assert.match(banner, /Date\.now\(\) - ts > maxAgeMs/);
  }
});

test("Universal Search keyboard actions use rendered result indexes", async () => {
  for (const file of ["popup.js", "SP-Developer-Toolkit-Lite/popup.js"]) {
    const popup = await text(file);
    assert.match(popup, /universalSearchVisibleItems/);
    assert.match(popup, /universalSearchVisibleItems\[universalSearchActive\]\.run\(\)/);
    assert.doesNotMatch(popup, /universalSearchFiltered\[universalSearchActive\]\.run\(\)/);
  }
});
