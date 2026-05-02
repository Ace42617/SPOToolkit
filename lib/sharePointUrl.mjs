export function isSharePointOnlineUrl(url, options = {}) {
  const { allowTenantAdmin = true } = options;
  try {
    const u = new URL(String(url || ""));
    if (u.protocol !== "https:") return false;
    const host = u.hostname.toLowerCase();
    if (!host.endsWith(".sharepoint.com")) return false;
    if (!allowTenantAdmin && host.endsWith("-admin.sharepoint.com")) return false;
    return true;
  } catch (_) {
    return false;
  }
}
