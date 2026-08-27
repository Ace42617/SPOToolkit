import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import {
  normalizeServerRelativePath,
  parentServerRelativePath,
  getInheritanceSourcePath,
  inheritedRoleAssignmentKind,
  inheritedRoleAssignmentUrl
} from "../lib/matrixInheritanceSource.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const LIST_ROOT = "/sites/legal/Shared Documents";
const CONTRACTS = "/sites/legal/Shared Documents/Contracts";
const NDA = "/sites/legal/Shared Documents/Contracts/NDA.docx";
const NESTED = "/sites/legal/Shared Documents/Contracts/2024/Q1/x.docx";
const ROOT_FILE = "/sites/legal/Shared Documents/readme.docx";

describe("normalizeServerRelativePath / parentServerRelativePath", () => {
  it("trims trailing slashes and duplicate separators", () => {
    assert.equal(normalizeServerRelativePath("/sites/legal/Shared Documents/"), LIST_ROOT);
    assert.equal(normalizeServerRelativePath("\\sites\\legal\\Shared Documents\\Contracts\\"), CONTRACTS);
  });

  it("returns the immediate parent folder", () => {
    assert.equal(parentServerRelativePath(NDA), CONTRACTS);
    assert.equal(parentServerRelativePath(CONTRACTS), LIST_ROOT);
    assert.equal(parentServerRelativePath("/sites"), "");
    assert.equal(parentServerRelativePath(""), "");
  });
});

describe("getInheritanceSourcePath", () => {
  it("uses a uniquely permissioned parent folder, not the list or web", () => {
    assert.equal(
      getInheritanceSourcePath(NDA, [CONTRACTS], LIST_ROOT),
      CONTRACTS
    );
  });

  it("picks the nearest unique ancestor when several folders break inheritance", () => {
    const unique = [
      CONTRACTS,
      "/sites/legal/Shared Documents/Contracts/2024"
    ];
    assert.equal(
      getInheritanceSourcePath(NESTED, unique, LIST_ROOT),
      "/sites/legal/Shared Documents/Contracts/2024"
    );
  });

  it("falls back to the list root when no unique folder sits above the item", () => {
    assert.equal(getInheritanceSourcePath(ROOT_FILE, [], LIST_ROOT), LIST_ROOT);
    assert.equal(getInheritanceSourcePath(NDA, [], LIST_ROOT), LIST_ROOT);
  });

  it("does not treat a uniquely permissioned sibling file as the source", () => {
    const unique = ["/sites/legal/Shared Documents/Contracts/secret.docx"];
    assert.equal(getInheritanceSourcePath(NDA, unique, LIST_ROOT), LIST_ROOT);
  });

  it("matches unique folder paths case-insensitively", () => {
    assert.equal(
      getInheritanceSourcePath(
        "/sites/legal/Shared Documents/contracts/NDA.docx",
        [CONTRACTS],
        LIST_ROOT
      ),
      CONTRACTS
    );
  });
});

describe("inheritedRoleAssignmentKind / URL", () => {
  const listBase = "https://contoso.sharepoint.com/sites/legal/_api/web/lists(guid'abc')";
  const webUrl = "https://contoso.sharepoint.com/sites/legal";

  it("reads the unique folder item RoleAssignments, not list or web", () => {
    assert.equal(
      inheritedRoleAssignmentKind(CONTRACTS, LIST_ROOT, true, 42),
      "item"
    );
    assert.equal(
      inheritedRoleAssignmentUrl("item", listBase, webUrl, 42),
      listBase + "/items(42)/RoleAssignments"
    );
  });

  it("reads list RoleAssignments when the list has unique perms and no unique folder", () => {
    assert.equal(
      inheritedRoleAssignmentKind(LIST_ROOT, LIST_ROOT, true, 0),
      "list"
    );
    assert.equal(
      inheritedRoleAssignmentUrl("list", listBase, webUrl, 0),
      listBase + "/roleassignments"
    );
  });

  it("reads web RoleAssignments when both the list and folders inherit", () => {
    assert.equal(
      inheritedRoleAssignmentKind(LIST_ROOT, LIST_ROOT, false, 0),
      "web"
    );
    assert.equal(
      inheritedRoleAssignmentUrl("web", listBase, webUrl, 0),
      webUrl + "/_api/web/roleassignments"
    );
  });

  it("does not use a list-root folder item id (PowerShell forces srcId=0 at list root)", () => {
    assert.equal(
      inheritedRoleAssignmentKind(LIST_ROOT, LIST_ROOT, false, 7),
      "web"
    );
    assert.equal(
      inheritedRoleAssignmentKind(LIST_ROOT, LIST_ROOT, true, 7),
      "list"
    );
  });
});

describe("permissionsMatrixExport.js stays in sync", () => {
  const source = readFileSync(join(root, "permissionsMatrixExport.js"), "utf8");

  it("inlines the inheritance helpers used by Include-all-inherited rows", () => {
    assert.match(source, /function getInheritanceSourcePath\s*\(/);
    assert.match(source, /function inheritedRoleAssignmentKind\s*\(/);
    assert.match(source, /function inheritedRoleAssignmentUrl\s*\(/);
    assert.match(
      source,
      /INCLUDE_ALL_INHERITED\) \{\s*var inhSrc = getInheritanceSourcePath\(/
    );
  });

  it("no longer assumes inherited files always use list or web role assignments", () => {
    assert.doesNotMatch(
      source,
      /else if \(!skipNonUniqueMatrix && INCLUDE_ALL_INHERITED\) \{\s*var inhUrl = listHasUnique \? listBase \+ "\/roleassignments" : webUrl \+ "\/_api\/web\/roleassignments"/
    );
  });
});
