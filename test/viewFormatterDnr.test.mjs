import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";

const root = new URL("..", import.meta.url);

describe("View Formatter iframe DNR rules", () => {
  it("does not ship global static rules for SharePoint subframes", () => {
    const manifest = JSON.parse(readFileSync(new URL("../manifest.json", import.meta.url), "utf8"));
    assert.ok(manifest.permissions.includes("declarativeNetRequest"));
    assert.equal(manifest.declarative_net_request, undefined);
    assert.equal(existsSync(join(root.pathname, "rules/view-formatter-iframe.json")), false);
  });

  it("scopes header stripping to requests initiated by this extension", () => {
    const background = readFileSync(new URL("../background.js", import.meta.url), "utf8");
    assert.match(background, /updateDynamicRules/);
    assert.match(background, /initiatorDomains:\s*\[initiator\]/);
    assert.match(background, /chrome\.runtime\s*&&\s*chrome\.runtime\.id/);
  });
});
