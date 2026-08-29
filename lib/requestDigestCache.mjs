/**
 * Request-digest cache helpers for Permissions Matrix export.
 * Keep in sync with permissionsMatrixExport.js (injected classic page script).
 */

/** Fallback TTL when SharePoint omits FormDigestTimeoutSeconds (~30m typical). */
export const DEFAULT_DIGEST_TTL_MS = 25 * 60 * 1000;

/** Refresh before the server-side timeout to avoid borderline 403s. */
export const DIGEST_TTL_SAFETY_RATIO = 0.8;

export function normalizeDigestCacheKey(webUrl) {
  return String(webUrl || "").toLowerCase();
}

/**
 * Build a cache entry from /_api/contextinfo JSON.
 * @returns {{ digest: string, expiresAt: number } | null}
 */
export function extractContextDigest(contextInfoJson, nowMs = Date.now()) {
  const info =
    (contextInfoJson &&
      contextInfoJson.d &&
      contextInfoJson.d.GetContextWebInformation) ||
    (contextInfoJson && contextInfoJson.GetContextWebInformation) ||
    contextInfoJson ||
    null;
  const digest = info && info.FormDigestValue;
  if (!digest) return null;
  const timeoutSec = Number(info.FormDigestTimeoutSeconds);
  const ttlMs =
    Number.isFinite(timeoutSec) && timeoutSec > 0
      ? Math.max(1000, Math.floor(timeoutSec * 1000 * DIGEST_TTL_SAFETY_RATIO))
      : DEFAULT_DIGEST_TTL_MS;
  return { digest: String(digest), expiresAt: nowMs + ttlMs };
}

/** Return a still-valid cached digest, or null if missing/expired. */
export function getCachedDigest(cache, webUrl, nowMs = Date.now()) {
  const key = normalizeDigestCacheKey(webUrl);
  const entry = cache && cache[key];
  if (!entry || !entry.digest) return null;
  if (typeof entry.expiresAt === "number" && nowMs >= entry.expiresAt) return null;
  return entry.digest;
}

export function setCachedDigest(cache, webUrl, digestEntry) {
  if (!cache || !digestEntry || !digestEntry.digest) return;
  cache[normalizeDigestCacheKey(webUrl)] = {
    digest: digestEntry.digest,
    expiresAt: digestEntry.expiresAt
  };
}

export function clearCachedDigest(cache, webUrl) {
  if (!cache) return;
  delete cache[normalizeDigestCacheKey(webUrl)];
}

/**
 * Whether a sharing GetSharingInformation response should trigger one digest refresh.
 * Only 403 is treated as a stale-digest signal; callers must pass alreadyRetried=true
 * after the first refresh so permanent ACL failures still fail closed to skip.
 */
export function shouldRetrySharingFetchAfterStatus(status, alreadyRetriedDigest) {
  return status === 403 && !alreadyRetriedDigest;
}
