import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { stripIdHashPrefix } from "../lib/sharePointLookupValue.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");

describe("stripIdHashPrefix", () => {
  it("strips classic ID;# lookup / user-value prefixes", () => {
    assert.equal(stripIdHashPrefix("12;#Contoso User"), "Contoso User");
    assert.equal(stripIdHashPrefix("0;#/sites/hr/Shared Documents/a.docx"), "/sites/hr/Shared Documents/a.docx");
    assert.equal(stripIdHashPrefix("5;#0x0120D52000"), "0x0120D52000");
  });

  it("preserves SharePoint paths and names that contain #", () => {
    assert.equal(
      stripIdHashPrefix("/sites/hr/Shared Documents/Budget#2024.xlsx"),
      "/sites/hr/Shared Documents/Budget#2024.xlsx"
    );
    assert.equal(stripIdHashPrefix("Budget#2024.xlsx"), "Budget#2024.xlsx");
    assert.equal(
      stripIdHashPrefix("/sites/hr/Shared Documents/Q1#Plan/Report.docx"),
      "/sites/hr/Shared Documents/Q1#Plan/Report.docx"
    );
  });

  it("leaves non-lookup strings and empty values alone", () => {
    assert.equal(stripIdHashPrefix(""), "");
    assert.equal(stripIdHashPrefix(null), "");
    assert.equal(stripIdHashPrefix("Document"), "Document");
    assert.equal(stripIdHashPrefix("#orphan"), "#orphan");
    assert.equal(stripIdHashPrefix("abc;#not-an-id-prefix"), "abc;#not-an-id-prefix");
  });
});

describe("export path strip sources stay in sync", () => {
  const sources = [
    "exportCSV.js",
    "permissionsMatrixExport.js"
  ].map((rel) => ({
    rel,
    source: readFileSync(join(root, rel), "utf8")
  }));

  for (const { rel, source } of sources) {
    it(`${rel} only strips leading digit;# lookup prefixes`, () => {
      assert.match(source, /function strip(?:IdHash|BagelPath)Prefix\s*\(/);
      assert.match(source, /\/\^\(\\d\+;#\)\//);
      // Former greedy first-# slice must not remain on path normalization paths.
      assert.doesNotMatch(
        source,
        /itemPath\.slice\(itemPath\.indexOf\("#"\)\s*\+\s*1\)/
      );
      assert.doesNotMatch(
        source,
        /var idx = s\.indexOf\("#"\);\s*return idx >= 0 \? s\.slice\(idx \+ 1\)/
      );
    });
  }

  it("permissions matrix item scan uses the safe strip helper", () => {
    const source = sources.find((e) => e.rel === "permissionsMatrixExport.js").source;
    assert.match(
      source,
      /itemPath\s*=\s*stripBagelPathPrefix\s*\(\s*itemPath\s*\)/
    );
  });
});
