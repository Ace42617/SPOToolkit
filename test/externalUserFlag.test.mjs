import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, it } from "node:test";
import { isExternalUser } from "../lib/externalUserFlag.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");

function read(rel) {
  return readFileSync(join(root, rel), "utf8");
}

describe("isExternalUser", () => {
  it("flags SharePoint guest logins with #ext#", () => {
    assert.equal(
      isExternalUser("i:0#.f|membership|alex_contoso.com#ext#@tenant.onmicrosoft.com", ""),
      true
    );
  });

  it("flags guest emails with #EXT#@", () => {
    assert.equal(isExternalUser("", "alex_contoso.com#EXT#@tenant.onmicrosoft.com"), true);
    assert.equal(isExternalUser("i:0#.f|membership|someone", "alex#ext#@tenant.onmicrosoft.com"), true);
  });

  it("does not flag Azure AD security groups (c:0t.c|tenant|)", () => {
    const aadGroup = "c:0t.c|tenant|6510e196-d412-41de-a2e3-f99e8c0ffb4a";
    assert.equal(isExternalUser(aadGroup, ""), false);
    assert.equal(isExternalUser(aadGroup.toUpperCase(), ""), false);
  });

  it("does not flag Microsoft 365 groups or Everyone-except-external", () => {
    assert.equal(
      isExternalUser("c:0o.c|federateddirectoryclaimprovider|45dfc99d-a4a5-419d-ad79-013eed517fb2", ""),
      false
    );
    assert.equal(isExternalUser("c:0-.f|rolemanager|spo-grid-all-users/tenant-id", ""), false);
  });

  it("does not flag ordinary member logins", () => {
    assert.equal(isExternalUser("i:0#.f|membership|alex@contoso.com", "alex@contoso.com"), false);
  });

  it("treats empty login and email as not external", () => {
    assert.equal(isExternalUser("", ""), false);
    assert.equal(isExternalUser(null, null), false);
  });
});

describe("source-sync: browser exporters must not treat c:0t.c as external", () => {
  const files = ["permissionsMatrixExport.js", "exportCSV.js"];

  for (const file of files) {
    it(`${file} guest check is #ext# / #EXT#@ only`, () => {
      const src = read(file);
      const fn = src.match(/function isExternalUser\([\s\S]*?\n  \}/);
      assert.ok(fn, `isExternalUser not found in ${file}`);
      assert.match(fn[0], /#ext#/i);
      assert.match(fn[0], /#EXT#@/i);
      assert.doesNotMatch(fn[0], /c:0t\.c/);
    });
  }
});

describe("source-sync: remediation must re-validate guest logins", () => {
  it("Build-ExternalUserPlan calls Test-ExternalLogin for matrix and group rows", () => {
    const src = read("scripts/Invoke-PermissionsMatrixRemediation.ps1");
    const start = src.indexOf("function Build-ExternalUserPlan");
    const next = src.indexOf("\nfunction Build-SharingLinkPlan", start + 1);
    assert.ok(start >= 0 && next > start, "Build-ExternalUserPlan not found");
    const body = src.slice(start, next);
    const matrixCalls = body.split("Test-ExternalLogin").length - 1;
    assert.ok(matrixCalls >= 2, "expected Test-ExternalLogin on matrix rows and group-member rows");
    assert.match(body, /Member Login/);
    assert.match(body, /Member Email/);
  });
});
