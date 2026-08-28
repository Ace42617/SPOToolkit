/**
 * Page SharePoint group membership. Keep in sync with the inlined copies in
 * permissionsMatrixExport.js (injected classic page scripts cannot import from
 * lib/) and Get-GroupMembers in Export-SitePermissionsMatrix.ps1.
 *
 * /sitegroups(...)/users defaults to 100 rows. SiteGroups/Users often omits
 * @odata.nextLink, so a full page must continue with $skip. Dedup by Id so an
 * ignored $skip cannot loop forever on the same first page.
 */

export const GROUP_USERS_PAGE_SIZE = 5000;
export const GROUP_USERS_MAX_PAGES = 50;
export const GROUP_USERS_SELECT = "Id,Title,LoginName,Email,PrincipalType";

export function buildGroupUsersPageUrl(usersCollectionUrl, pageSize, skip) {
  var base = String(usersCollectionUrl || "");
  var size = pageSize > 0 ? pageSize : GROUP_USERS_PAGE_SIZE;
  var q = "$select=" + GROUP_USERS_SELECT + "&$top=" + size;
  if (skip > 0) q += "&$skip=" + skip;
  var sep = base.indexOf("?") >= 0 ? "&" : "?";
  return base + sep + q;
}

export function groupUsersODataNextLink(payload) {
  if (!payload || typeof payload !== "object") return null;
  var next = payload["@odata.nextLink"] || payload["odata.nextLink"] || null;
  if (!next && payload.d) next = payload.d.__next || payload.d["odata.nextLink"] || null;
  if (next == null) return null;
  var s = String(next).trim();
  return s || null;
}

export function groupUsersPageItems(payload) {
  if (!payload) return [];
  if (Array.isArray(payload.value)) return payload.value;
  if (Array.isArray(payload.results)) return payload.results;
  if (payload.d) {
    if (Array.isArray(payload.d.results)) return payload.d.results;
    if (Array.isArray(payload.d.value)) return payload.d.value;
  }
  return [];
}

export function groupUserKey(user) {
  if (!user || typeof user !== "object") return "";
  if (user.Id != null && String(user.Id).trim() !== "") return "id:" + String(user.Id).trim();
  var login = String(user.LoginName || user.loginName || "").trim();
  if (login) return "login:" + login.toLowerCase();
  return "";
}

/**
 * After a page: stop, follow nextLink, or $skip from receivedCount + pageLength.
 * addedCount === 0 means empty or a duplicate page ($skip ignored).
 */
export function groupUsersPagingAction(opts) {
  opts = opts || {};
  var addedCount = opts.addedCount;
  var pageLength = opts.pageLength || 0;
  var pageSize = opts.pageSize > 0 ? opts.pageSize : GROUP_USERS_PAGE_SIZE;
  var nextLink = opts.nextLink || null;
  if (!addedCount) return { action: "stop" };
  if (nextLink) return { action: "nextLink", url: nextLink };
  if (pageLength >= pageSize) return { action: "skip" };
  return { action: "stop" };
}

export function mergeUniqueGroupUsers(users, page, seen) {
  var added = 0;
  var list = users || [];
  var keys = seen || Object.create(null);
  var rows = page || [];
  for (var i = 0; i < rows.length; i++) {
    var u = rows[i];
    var key = groupUserKey(u);
    if (key && keys[key]) continue;
    if (key) keys[key] = true;
    list.push(u);
    added++;
  }
  return { users: list, added: added, seen: keys };
}

/**
 * Drive paging with a fetchPage(url) that returns a REST payload or null.
 */
export async function collectGroupUsersFromPages(fetchPage, usersCollectionUrl, pageSize) {
  var size = pageSize > 0 ? pageSize : GROUP_USERS_PAGE_SIZE;
  var users = [];
  var seen = Object.create(null);
  var receivedCount = 0;
  var url = buildGroupUsersPageUrl(usersCollectionUrl, size, 0);
  var pages = 0;
  while (url && pages < GROUP_USERS_MAX_PAGES) {
    pages++;
    var payload = await fetchPage(url);
    if (!payload) break;
    var page = groupUsersPageItems(payload);
    var merged = mergeUniqueGroupUsers(users, page, seen);
    users = merged.users;
    seen = merged.seen;
    var action = groupUsersPagingAction({
      addedCount: merged.added,
      pageLength: page.length,
      pageSize: size,
      nextLink: groupUsersODataNextLink(payload)
    });
    if (action.action === "stop") break;
    if (action.action === "nextLink") {
      url = action.url;
      receivedCount += page.length;
      continue;
    }
    receivedCount += page.length;
    url = buildGroupUsersPageUrl(usersCollectionUrl, size, receivedCount);
  }
  return users;
}
