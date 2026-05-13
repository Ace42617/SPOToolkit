import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const root = new URL("../", import.meta.url);

describe("View Formatter frame header rules", () => {
  it("does not enable browser-wide static DNR rules", async () => {
    const manifest = JSON.parse(await readFile(new URL("manifest.json", root), "utf8"));
    assert.equal(manifest.declarative_net_request, undefined);
  });

  it("scopes formatter frame rules to session tab IDs", async () => {
    const background = await readFile(new URL("background.js", root), "utf8");
    assert.match(background, /updateSessionRules/);
    assert.match(background, /tabIds/);
    assert.match(background, /SPOToolkitOpenViewFormatter/);
  });

  it("routes popup formatter launches through the background worker", async () => {
    const popup = await readFile(new URL("popup.js", root), "utf8");
    assert.match(popup, /SPOToolkitOpenViewFormatter/);
    assert.doesNotMatch(popup, /chrome\.tabs\.create\(\{\s*url:\s*url\.toString\(\)\s*\}\)/);
  });
});
