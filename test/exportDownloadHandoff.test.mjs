import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, it } from "node:test";
import {
  hasExportDownloadRequestId,
  unwrapPageExportDetail
} from "../lib/exportDownloadHandoff.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");

describe("export download handoff", () => {
  it("unwraps page postMessage detail once and preserves requestId", () => {
    const pageMsg = {
      __spcsv: true,
      type: "SPCSVExportDownloadRequest",
      detail: {
        requestId: "dl-1-123",
        path: "Export.zip",
        mime: "application/zip",
        text: "x"
      }
    };
    const d = unwrapPageExportDetail(pageMsg);
    assert.equal(d.requestId, "dl-1-123");
    assert.equal(hasExportDownloadRequestId(d), true);
    // Double-unwrap (the v1.1.0.0 bug) loses requestId.
    const doubleUnwrapped = d.detail || {};
    assert.equal(hasExportDownloadRequestId(doubleUnwrapped), false);
  });

  it("preserves chunk and chunk-end requestIds after one unwrap", () => {
    const chunk = unwrapPageExportDetail({
      __spcsv: true,
      type: "SPCSVExportDownloadChunk",
      detail: {
        requestId: "dl-2-456",
        path: "a.zip",
        mime: "application/zip",
        chunkIndex: 0,
        totalChunks: 1,
        totalBytes: 4,
        buffer: new ArrayBuffer(4)
      }
    });
    const end = unwrapPageExportDetail({
      __spcsv: true,
      type: "SPCSVExportDownloadChunkEnd",
      detail: { requestId: "dl-2-456" }
    });
    assert.equal(chunk.requestId, "dl-2-456");
    assert.equal(end.requestId, "dl-2-456");
  });

  it("content.js download handlers use the already-unwrapped detail (source sync)", () => {
    const src = readFileSync(join(root, "content.js"), "utf8");
    assert.match(
      src,
      /SPCSVExportDownloadRequest"\)\s*\{\s*(?:\/\/[^\n]*\n\s*)*void relayExportDownloadRequest\(d\);/
    );
    assert.match(
      src,
      /SPCSVExportDownloadChunk"\)\s*\{\s*void handleExportDownloadChunk\(d\);/
    );
    assert.match(
      src,
      /SPCSVExportDownloadChunkEnd"\)\s*\{\s*void finalizeExportDownloadTransfer\(d\);/
    );
    assert.doesNotMatch(
      src,
      /relayExportDownloadRequest\(d\.detail/
    );
    assert.doesNotMatch(
      src,
      /handleExportDownloadChunk\(d\.detail/
    );
    assert.doesNotMatch(
      src,
      /finalizeExportDownloadTransfer\(d\.detail/
    );
  });

  it("background buffer download waits for completion before resolving (source sync)", () => {
    const src = readFileSync(join(root, "background.js"), "utf8");
    const fnStart = src.indexOf("function downloadBufferInBackground");
    assert.ok(fnStart >= 0, "downloadBufferInBackground present");
    const fnEnd = src.indexOf("\nfunction ", fnStart + 1);
    const body = src.slice(fnStart, fnEnd > fnStart ? fnEnd : undefined);
    // Must not resolve() before waitForBackgroundDownload settles.
    assert.doesNotMatch(
      body,
      /resolve\(\);\s*waitForBackgroundDownload/
    );
    assert.match(
      body,
      /waitForBackgroundDownload\(downloadId,\s*timeoutMs\)\s*\.then\(/
    );
    assert.match(body, /reject\(err/);
  });
});
