import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { join } from "node:path";

const root = process.cwd();
const packages = [
  {
    name: "root",
    manifest: "manifest.json",
    rules: "rules/view-formatter-iframe.json",
    background: "background.js",
  },
  {
    name: "experimental",
    manifest: "SP-Developer-Toolkit-Experimental/manifest.json",
    rules: "SP-Developer-Toolkit-Experimental/rules/view-formatter-iframe.json",
    background: "SP-Developer-Toolkit-Experimental/background.js",
  },
];

function readJson(relativePath) {
  return JSON.parse(readFileSync(join(root, relativePath), "utf8"));
}

function readText(relativePath) {
  return readFileSync(join(root, relativePath), "utf8");
}

function stripsSharePointFrameProtection(rule) {
  const headers = rule?.action?.responseHeaders || [];
  return (
    rule?.action?.type === "modifyHeaders" &&
    headers.some((h) => h.operation === "remove" && String(h.header || "").toLowerCase() === "x-frame-options") &&
    headers.some((h) => h.operation === "remove" && String(h.header || "").toLowerCase() === "content-security-policy")
  );
}

describe("View Formatter frame DNR rules", () => {
  it("does not enable static SharePoint frame-protection stripping rulesets", () => {
    for (const pkg of packages) {
      const manifest = readJson(pkg.manifest);
      const resources = manifest.declarative_net_request?.rule_resources || [];
      const formatterRuleset = resources.find((resource) => resource.path === "rules/view-formatter-iframe.json");
      assert.ok(formatterRuleset, `${pkg.name} manifest should declare the formatter ruleset`);
      assert.equal(formatterRuleset.enabled, false, `${pkg.name} formatter ruleset must be disabled by default`);

      const rules = readJson(pkg.rules);
      assert.ok(
        rules.some(stripsSharePointFrameProtection),
        `${pkg.name} formatter rules should be the only disabled static copy of the header-stripping rules`
      );
    }
  });

  it("installs formatter frame rules only as tab- and initiator-scoped session rules", () => {
    for (const pkg of packages) {
      const background = readText(pkg.background);
      assert.match(background, /updateSessionRules/, `${pkg.name} background must use session DNR rules`);
      assert.match(background, /tabIds/, `${pkg.name} session rules must be scoped to formatter tab ids`);
      assert.match(background, /initiatorDomains\s*=\s*\[chrome\.runtime\.id\]/, `${pkg.name} rules must be scoped to this extension`);
      assert.match(
        background,
        /SPOToolkitEnsureViewFormatterFrameRules/,
        `${pkg.name} formatter page must be able to re-register its scoped rules`
      );
    }
  });
});
