import assert from "node:assert/strict";
import fs from "node:fs";
import test from "node:test";

const bundle = fs.readFileSync(new URL("../levelup/background.js", import.meta.url), "utf8");
const vendor = fs.readFileSync(
  new URL("../vendor/levelup/src/background/background.ts", import.meta.url),
  "utf8"
);
const overlay = fs.readFileSync(
  new URL("../scripts/levelup-adaptations/overlay/src/background/background.ts", import.meta.url),
  "utf8"
);

test("impersonation rules are not wiped on every service worker start", () => {
  assert.equal(bundle.includes("(async()=>{await a.initializeOnStartup()})()"), false);
  assert.match(bundle, /onInstalled\.addListener\(async\(\)=>\{await a\.initializeOnStartup\(\)\}\)/);
  assert.match(bundle, /onStartup\.addListener\(async\(\)=>\{await a\.initializeOnStartup\(\)\}\)/);

  for (const source of [vendor, overlay]) {
    assert.match(source, /chrome\.runtime\.onInstalled\.addListener/);
    assert.match(source, /chrome\.runtime\.onStartup\.addListener/);
    assert.equal(source.includes("await impersonationService.initializeOnStartup();"), true);
    assert.doesNotMatch(source, /initializeOnStartup\(\);\s*\}\)\(\);/);
  }
});
