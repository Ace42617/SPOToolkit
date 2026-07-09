import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";

const REPO_ROOT = path.resolve(new URL("..", import.meta.url).pathname);
const BUILDS = [
  { name: "production", dir: "" },
  { name: "experimental", dir: "SP-Developer-Toolkit-Experimental" },
];

async function readText(build, file) {
  return readFile(path.join(REPO_ROOT, build.dir, file), "utf8");
}

describe("View Formatter iframe header rules", () => {
  for (const build of BUILDS) {
    it(`${build.name} manifest does not install static global DNR rules`, async () => {
      const manifest = JSON.parse(await readText(build, "manifest.json"));

      assert.equal(
        Object.hasOwn(manifest, "declarative_net_request"),
        false,
        "SharePoint frame header stripping must not be installed as a global static ruleset"
      );
      assert.ok(
        manifest.permissions.includes("declarativeNetRequest"),
        "the formatter still needs DNR permission for scoped session rules"
      );
    });

    it(`${build.name} formatter rules are registered per formatter tab`, async () => {
      const background = await readText(build, "background.js");
      const formatter = await readText(build, "view-formatter.js");

      assert.match(background, /updateSessionRules/);
      assert.match(background, /tabIds/);
      assert.match(background, /SPOToolkitRegisterFormatterPreviewTab/);
      assert.match(formatter, /SPOToolkitRegisterFormatterPreviewTab/);
      assert.match(formatter, /pendingPreviewFrameUrl/);
    });
  }
});
