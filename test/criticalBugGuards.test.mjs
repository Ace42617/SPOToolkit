import { describe, it } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";

const root = path.resolve(new URL("..", import.meta.url).pathname);

function read(rel) {
  return fs.readFileSync(path.join(root, rel), "utf8");
}

describe("critical remediation guards", () => {
  it("never removes sharing links without a concrete ShareId identity", () => {
    const script = read("scripts/Invoke-PermissionsMatrixRemediation.ps1");
    const removalLines = script
      .split(/\r?\n/)
      .filter((line) => /Remove-PnP(?:File|Folder)SharingLink/.test(line));

    assert.ok(removalLines.length > 0, "expected sharing-link removal commands");
    assert.deepEqual(
      removalLines.filter((line) => !line.includes("-Identity $shareId")),
      [],
      "sharing-link removal must always pass the selected ShareId"
    );
    assert.match(script, /if \(-not \$shareId\) \{\s*throw "Sharing link identity was missing/s);
  });

  it("does not synthesize parent-scope cleanup from file external-user rows", () => {
    const script = read("scripts/Invoke-PermissionsMatrixRemediation.ps1");
    assert.equal(script.includes("Remove site-level grants for external user"), false);
    assert.equal(script.includes("Remove library-level grants for external user"), false);
  });
});

describe("View Manager field-save guard", () => {
  for (const rel of ["views.js", "SP-Developer-Toolkit-Experimental/views.js"]) {
    it(`${rel} does not clear all ViewFields before re-adding columns`, () => {
      const source = read(rel);
      assert.equal(source.includes("RemoveAllViewFields"), false);
      assert.match(source, /addviewfield\('/);
      assert.match(source, /removeviewfield\('/);
      assert.match(source, /moveviewfieldto\(field='/);
    });
  }
});

describe("View Formatter frame-rule scoping", () => {
  for (const rel of ["manifest.json", "SP-Developer-Toolkit-Experimental/manifest.json"]) {
    it(`${rel} has no static DNR ruleset`, () => {
      const manifest = JSON.parse(read(rel));
      assert.equal(Object.hasOwn(manifest, "declarative_net_request"), false);
    });
  }

  for (const rel of ["background.js", "SP-Developer-Toolkit-Experimental/background.js"]) {
    it(`${rel} installs DNR rules scoped to the formatter tab`, () => {
      const source = read(rel);
      assert.match(source, /SPOToolkitInstallViewFormatterFrameRules/);
      assert.match(source, /tabIds: \[tabId\]/);
      assert.match(source, /updateSessionRules/);
    });
  }
});
