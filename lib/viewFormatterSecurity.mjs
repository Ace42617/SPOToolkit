export function normalizeSharePointPreviewUrl(rawUrl) {
  const value = String(rawUrl || "").trim();
  if (!value) return "";

  let parsed;
  try {
    parsed = new URL(value);
  } catch (_) {
    return "";
  }

  if (parsed.protocol !== "https:") return "";
  if (parsed.username || parsed.password) return "";

  const hostname = parsed.hostname.toLowerCase();
  if (!hostname.endsWith(".sharepoint.com")) return "";
  if (hostname.endsWith("-admin.sharepoint.com")) return "";

  return parsed.href;
}

export function isSharePointPreviewUrl(rawUrl) {
  return normalizeSharePointPreviewUrl(rawUrl) !== "";
}
