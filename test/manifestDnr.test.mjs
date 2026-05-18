import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const manifest = JSON.parse(await readFile(new URL("../manifest.json", import.meta.url), "utf8"));

describe("manifest DNR scoping", () => {
  it("does not ship global static SharePoint iframe header-stripping rules", () => {
    assert.equal(manifest.declarative_net_request, undefined);
  });

  it("keeps DNR permission for formatter-tab-scoped session rules", () => {
    assert.ok(manifest.permissions.includes("declarativeNetRequest"));
  });
});
