import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { existsSync, readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const testDir = dirname(fileURLToPath(import.meta.url));
const root = join(testDir, "..");

describe("View Formatter iframe DNR rules", () => {
  it("does not ship global static rules for SharePoint subframes", () => {
    const manifest = JSON.parse(readFileSync(join(root, "manifest.json"), "utf8"));
    assert.ok(manifest.permissions.includes("declarativeNetRequest"));
    assert.equal(manifest.declarative_net_request, undefined);
    assert.equal(existsSync(join(root, "rules", "view-formatter-iframe.json")), false);
  });

  it("installs header stripping only for formatter preview tabs", () => {
    const background = readFileSync(join(root, "background.js"), "utf8");
    assert.match(background, /updateSessionRules/);
    assert.match(background, /tabIds:\s*\[tabId\]/);
    assert.match(background, /resourceTypes:\s*\["sub_frame"\]/);
    assert.match(background, /chrome\.tabs\.onRemoved\.addListener/);
  });

  it("routes popup View Formatter opens through the scoped background helper", () => {
    const popup = readFileSync(join(root, "popup.js"), "utf8");
    assert.doesNotMatch(popup, /chrome\.tabs\.create\(\{\s*url:\s*url\.toString\(\)\s*\}\)/);
    assert.match(popup, /SPOToolkitOpenViewFormatter/);
  });
});
