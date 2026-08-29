import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import {
  DEFAULT_DIGEST_TTL_MS,
  DIGEST_TTL_SAFETY_RATIO,
  clearCachedDigest,
  extractContextDigest,
  getCachedDigest,
  setCachedDigest,
  shouldRetrySharingFetchAfterStatus
} from "../lib/requestDigestCache.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const exportSource = readFileSync(join(root, "permissionsMatrixExport.js"), "utf8");

describe("extractContextDigest", () => {
  it("uses FormDigestTimeoutSeconds with a safety margin", () => {
    const now = 1_000_000;
    const entry = extractContextDigest(
      {
        d: {
          GetContextWebInformation: {
            FormDigestValue: "abc",
            FormDigestTimeoutSeconds: 1800
          }
        }
      },
      now
    );
    assert.equal(entry.digest, "abc");
    assert.equal(
      entry.expiresAt,
      now + Math.floor(1800 * 1000 * DIGEST_TTL_SAFETY_RATIO)
    );
  });

  it("falls back to DEFAULT_DIGEST_TTL_MS when timeout is missing", () => {
    const now = 5_000;
    const entry = extractContextDigest(
      { d: { GetContextWebInformation: { FormDigestValue: "xyz" } } },
      now
    );
    assert.equal(entry.digest, "xyz");
    assert.equal(entry.expiresAt, now + DEFAULT_DIGEST_TTL_MS);
  });

  it("returns null without FormDigestValue", () => {
    assert.equal(extractContextDigest({ d: { GetContextWebInformation: {} } }), null);
  });
});

describe("getCachedDigest / setCachedDigest / clearCachedDigest", () => {
  it("returns cached digests until expiry, then null", () => {
    const cache = {};
    const now = 10_000;
    const web = "https://contoso.sharepoint.com/sites/a";
    setCachedDigest(cache, web, {
      digest: "d1",
      expiresAt: now + 1000
    });
    assert.equal(getCachedDigest(cache, web.toUpperCase(), now), "d1");
    assert.equal(getCachedDigest(cache, web, now + 1000), null);
  });

  it("clearCachedDigest removes a web entry so the next fetch refreshes", () => {
    const cache = {};
    setCachedDigest(cache, "https://contoso.sharepoint.com/sites/b", {
      digest: "old",
      expiresAt: Date.now() + 60_000
    });
    clearCachedDigest(cache, "https://contoso.sharepoint.com/sites/b");
    assert.equal(
      getCachedDigest(cache, "https://contoso.sharepoint.com/sites/b"),
      null
    );
  });
});

describe("shouldRetrySharingFetchAfterStatus", () => {
  it("retries only the first 403 (stale digest), not permanent ACL failures", () => {
    assert.equal(shouldRetrySharingFetchAfterStatus(403, false), true);
    assert.equal(shouldRetrySharingFetchAfterStatus(403, true), false);
    assert.equal(shouldRetrySharingFetchAfterStatus(404, false), false);
    assert.equal(shouldRetrySharingFetchAfterStatus(400, false), false);
  });
});

describe("permissionsMatrixExport.js stays in sync", () => {
  it("caches digest objects with expiresAt and refreshes on 403", () => {
    assert.match(exportSource, /DEFAULT_DIGEST_TTL_MS\s*=\s*25\s*\*\s*60\s*\*\s*1000/);
    assert.match(exportSource, /DIGEST_TTL_SAFETY_RATIO\s*=\s*0\.8/);
    assert.match(exportSource, /FormDigestTimeoutSeconds/);
    assert.match(exportSource, /expiresAt/);
    assert.match(exportSource, /async function getRequestDigest\(webUrl,\s*forceRefresh\)/);
    assert.match(exportSource, /var digestRetried\s*=\s*false/);
    assert.match(exportSource, /saw403/);
    assert.match(
      exportSource,
      /Stale FormDigest often surfaces as 403; refresh once before permanent skip/
    );
    // Must not cache bare digest strings forever.
    assert.doesNotMatch(
      exportSource,
      /requestDigestCache\[key\]\s*=\s*digest\s*;/
    );
  });
});
