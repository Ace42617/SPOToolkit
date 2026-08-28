import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import {
  GROUP_USERS_PAGE_SIZE,
  GROUP_USERS_SELECT,
  buildGroupUsersPageUrl,
  groupUsersODataNextLink,
  groupUsersPageItems,
  groupUserKey,
  groupUsersPagingAction,
  mergeUniqueGroupUsers,
  collectGroupUsersFromPages
} from "../lib/groupUsersPaging.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const collection = "https://contoso.sharepoint.com/sites/legal/_api/web/sitegroups/GetById(5)/users";

function users(from, count) {
  const rows = [];
  for (let i = 0; i < count; i++) {
    const id = from + i;
    rows.push({ Id: id, Title: "User " + id, LoginName: "i:0#.f|membership|u" + id + "@contoso.com", Email: "u" + id + "@contoso.com", PrincipalType: 1 });
  }
  return rows;
}

describe("buildGroupUsersPageUrl", () => {
  it("requests $top=5000 so the first page is not the REST default of 100", () => {
    const url = buildGroupUsersPageUrl(collection, GROUP_USERS_PAGE_SIZE, 0);
    assert.ok(url.includes("$select=" + GROUP_USERS_SELECT));
    assert.match(url, /\$top=5000/);
    assert.doesNotMatch(url, /\$skip=/);
  });

  it("adds $skip when continuing a full page that had no nextLink", () => {
    const url = buildGroupUsersPageUrl(collection, 5000, 5000);
    assert.match(url, /\$top=5000/);
    assert.match(url, /\$skip=5000/);
  });
});

describe("groupUsersPageItems / nextLink", () => {
  it("reads nometadata value arrays and @odata.nextLink", () => {
    const payload = { value: users(1, 2), "@odata.nextLink": collection + "?$skiptoken=x" };
    assert.equal(groupUsersPageItems(payload).length, 2);
    assert.equal(groupUsersODataNextLink(payload), collection + "?$skiptoken=x");
  });

  it("reads verbose d.results and d.__next", () => {
    const payload = { d: { results: users(1, 1), __next: collection + "?$skiptoken=v" } };
    assert.equal(groupUsersPageItems(payload).length, 1);
    assert.equal(groupUsersODataNextLink(payload), collection + "?$skiptoken=v");
  });
});

describe("groupUsersPagingAction", () => {
  it("stops on a short page", () => {
    assert.deepEqual(
      groupUsersPagingAction({ addedCount: 40, pageLength: 40, pageSize: 100, nextLink: null }),
      { action: "stop" }
    );
  });

  it("follows nextLink when SharePoint provides one", () => {
    assert.deepEqual(
      groupUsersPagingAction({ addedCount: 100, pageLength: 100, pageSize: 100, nextLink: "https://next" }),
      { action: "nextLink", url: "https://next" }
    );
  });

  it("falls back to $skip when a full page has no nextLink (SiteGroups/Users quirk)", () => {
    assert.deepEqual(
      groupUsersPagingAction({ addedCount: 100, pageLength: 100, pageSize: 100, nextLink: null }),
      { action: "skip" }
    );
  });

  it("stops when a page adds no new members so ignored $skip cannot loop", () => {
    assert.deepEqual(
      groupUsersPagingAction({ addedCount: 0, pageLength: 100, pageSize: 100, nextLink: null }),
      { action: "stop" }
    );
  });
});

describe("mergeUniqueGroupUsers", () => {
  it("dedupes by Id so a repeated first page is not appended again", () => {
    const page = users(1, 3);
    const first = mergeUniqueGroupUsers([], page, Object.create(null));
    const second = mergeUniqueGroupUsers(first.users, page, first.seen);
    assert.equal(first.added, 3);
    assert.equal(second.added, 0);
    assert.equal(second.users.length, 3);
  });

  it("uses LoginName when Id is missing", () => {
    assert.equal(groupUserKey({ LoginName: "i:0#.f|membership|A@contoso.com" }), "login:i:0#.f|membership|a@contoso.com");
  });
});

describe("collectGroupUsersFromPages", () => {
  it("does not stop at the REST default of 100 when more members exist", async () => {
    const page1 = users(1, 100);
    const page2 = users(101, 50);
    const byUrl = Object.create(null);
    byUrl[buildGroupUsersPageUrl(collection, 100, 0)] = { value: page1 };
    byUrl[buildGroupUsersPageUrl(collection, 100, 100)] = { value: page2 };

    const got = await collectGroupUsersFromPages(async (url) => byUrl[url] || null, collection, 100);
    assert.equal(got.length, 150);
    assert.equal(got[0].Id, 1);
    assert.equal(got[99].Id, 100);
    assert.equal(got[100].Id, 101);
    assert.equal(got[149].Id, 150);
  });

  it("follows @odata.nextLink across pages", async () => {
    const next = collection + "?$skiptoken=page2";
    const got = await collectGroupUsersFromPages(async (url) => {
      if (url.indexOf("$skiptoken=page2") >= 0) return { value: users(101, 20) };
      return { value: users(1, 100), "@odata.nextLink": next };
    }, collection, 100);
    assert.equal(got.length, 120);
  });

  it("stops when $skip is ignored and the first page is returned again", async () => {
    const same = { value: users(1, 100) };
    const got = await collectGroupUsersFromPages(async () => same, collection, 100);
    assert.equal(got.length, 100);
  });
});

describe("permissionsMatrixExport.js stays in sync", () => {
  const source = readFileSync(join(root, "permissionsMatrixExport.js"), "utf8");

  it("pages group members with $top=5000, nextLink, and $skip", () => {
    assert.match(source, /async function fetchGroupUsers\s*\(/);
    assert.match(source, /sitegroups\/GetById\(" \+ groupId \+ "\)\/users\?\$select=Id,Title,LoginName,Email,PrincipalType&\$top=/);
    assert.match(source, /\$skip=/);
    assert.match(source, /@odata\.nextLink/);
    assert.match(source, /odata\.nextLink/);
  });

  it("no longer fetches group members in a single unpaged request", () => {
    assert.doesNotMatch(
      source,
      /GetById\(" \+ groupId \+ "\)\/users\?\$select=Id,Title,LoginName,Email,PrincipalType"\)/
    );
  });
});

describe("Export-SitePermissionsMatrix.ps1 stays in sync", () => {
  const source = readFileSync(join(root, "scripts/Export-SitePermissionsMatrix.ps1"), "utf8");

  it("pages Get-GroupMembers with $top, nextLink, and $skip", () => {
    assert.match(source, /function Get-GroupMembers\s*\(\[int\]\s*\$GroupId\)[\s\S]*?Get-ODataNextLink \$r/);
    assert.match(source, /sitegroups\(\$GroupId\)\/users\?`\$select=Id,Title,LoginName,Email,PrincipalType&`\$top=/);
    assert.match(source, /function Get-GroupMembers[\s\S]*?`\$skip=/);
  });

  it("no longer fetches group members in a single unpaged request", () => {
    assert.doesNotMatch(
      source,
      /sitegroups\(\$GroupId\)\/users\?`\$select=Id,Title,LoginName,Email,PrincipalType"\s*\n/
    );
  });
});

void GROUP_USERS_SELECT;
