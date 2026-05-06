/**
 * Security regression tests for SharePoint iframe header handling.
 */
import { describe, it } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";

const root = path.resolve(new URL("..", import.meta.url).pathname);

describe("declarativeNetRequest iframe scope", () => {
  it("does not register static global SharePoint frame header-stripping rules", () => {
    const manifest = JSON.parse(fs.readFileSync(path.join(root, "manifest.json"), "utf8"));
    const ruleResources = manifest.declarative_net_request?.rule_resources || [];
    assert.equal(ruleResources.length, 0);

    const staticRules = JSON.parse(fs.readFileSync(path.join(root, "rules/view-formatter-iframe.json"), "utf8"));
    assert.deepEqual(staticRules, []);
  });
});
