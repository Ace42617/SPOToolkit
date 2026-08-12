import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import {
  normalizeServerRelativePath,
  resolvePathLengthLibraryRoot,
  pathRelativeToLibraryRoot
} from "../lib/pathLengthLibraryRoot.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");

describe("resolvePathLengthLibraryRoot", () => {
  it("prefers explicit library RootFolder over shortest file parent", () => {
    const listRoot = "/sites/contoso/Shared Documents";
    // Shortest file lives under A/; old code would treat A as the library root.
    const parents = [
      "/sites/contoso/Shared Documents",
      "/sites/contoso/Shared Documents/A"
    ];
    assert.equal(resolvePathLengthLibraryRoot(listRoot, parents), listRoot);
    assert.equal(
      resolvePathLengthLibraryRoot(listRoot, ["/sites/contoso/Shared Documents/A"]),
      listRoot
    );
  });

  it("falls back to shortest parent/FileDirRef when RootFolder is missing", () => {
    assert.equal(
      resolvePathLengthLibraryRoot("", [
        "/sites/contoso/Shared Documents/A",
        "/sites/contoso/Shared Documents"
      ]),
      "/sites/contoso/Shared Documents"
    );
  });
});

describe("pathRelativeToLibraryRoot", () => {
  const listRoot = "/sites/contoso/Shared Documents";

  it("keeps library-relative paths for root and nested files", () => {
    assert.equal(
      pathRelativeToLibraryRoot(
        "/sites/contoso/Shared Documents/Annual-Financial-Report-Consolidated-FY2024-Final.docx",
        listRoot
      ),
      "Annual-Financial-Report-Consolidated-FY2024-Final.docx"
    );
    assert.equal(
      pathRelativeToLibraryRoot("/sites/contoso/Shared Documents/A/x.docx", listRoot),
      "A/x.docx"
    );
  });

  it("does not undercount nested paths when a short branch exists", () => {
    // Reproduce the old shortest-file-parent bug: root inferred as .../A
    const wrongRoot = "/sites/contoso/Shared Documents/A";
    assert.equal(
      pathRelativeToLibraryRoot("/sites/contoso/Shared Documents/A/x.docx", wrongRoot),
      "x.docx"
    );
    assert.notEqual(
      pathRelativeToLibraryRoot("/sites/contoso/Shared Documents/A/x.docx", listRoot),
      "x.docx"
    );
  });

  it("normalizes trailing slashes", () => {
    assert.equal(normalizeServerRelativePath("/sites/x/Lib/"), "/sites/x/Lib");
    assert.equal(
      pathRelativeToLibraryRoot("/sites/contoso/Shared Documents/", listRoot + "/"),
      ""
    );
  });
});

describe("path length sources stay in sync", () => {
  const sources = ["exportCSV.js", "permissionsMatrixExport.js"].map((rel) => ({
    rel,
    source: readFileSync(join(root, rel), "utf8")
  }));

  for (const { rel, source } of sources) {
    it(`${rel} does not infer Path Lengths root from the shortest file alone`, () => {
      assert.doesNotMatch(
        source,
        /shortestPath\s*=\s*null[\s\S]{0,220}?normRoot\s*=\s*\(shortestPath\s*\?/
      );
    });

    it(`${rel} resolves Path Lengths root via library root helper pattern`, () => {
      assert.match(source, /function resolvePathLengthLibraryRoot\s*\(/);
      assert.match(source, /function pathRelativeToLibraryRoot\s*\(/);
    });
  }

  it("bagel Path Lengths passes list RootFolder into appendBagelPathLengthRows", () => {
    const source = sources.find((e) => e.rel === "permissionsMatrixExport.js").source;
    assert.match(
      source,
      /appendBagelPathLengthRows\s*\(\s*bagelPathRows\s*,\s*allItems\s*,\s*siteName\s*,\s*listTitle\s*,\s*listRootUrl\s*\)/
    );
  });

  it("standalone Path Lengths passes resolved library root into the builder", () => {
    const source = sources.find((e) => e.rel === "exportCSV.js").source;
    assert.match(
      source,
      /buildPathLengthsReportFromRows\s*\(\s*[^,]+,\s*(?:await\s+)?resolveListRootServerRelativeUrl\s*\(\s*\)\s*\)/
    );
  });
});
