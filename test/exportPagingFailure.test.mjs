import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const variants = [
  "../exportCSV.js",
  "../SP-Developer-Toolkit-Experimental/exportCSV.js",
  "../SP-Developer-Toolkit-Lite/exportCSV.js",
];

for (const relativePath of variants) {
  const source = await readFile(new URL(relativePath, import.meta.url), "utf8");
  const label = relativePath.replace("../", "");
  const owssvrStart = source.indexOf("async function exportViaOwssvrWithView");
  const restStart = source.indexOf("async function exportViaRest", owssvrStart);
  const owssvrSource = source.slice(owssvrStart, restStart);
  const restSource = source.slice(restStart);
  const helperSource = source.match(
    /function owssvrIncompleteMessage\([^)]*\) \{[\s\S]*?\n  \}/
  );

  describe(`${label} paging failures`, () => {
    it("labels partial OWSSVR results as incomplete", () => {
      assert.ok(helperSource, `${label} should define the incomplete-export message helper`);
      const messageFor = Function(`return (${helperSource[0]});`)();
      const message = messageFor("owssvr.dll HTTP 503: Service Unavailable.", 2_500);

      assert.match(message, /Export incomplete after 2,500 rows/);
      assert.match(message, /No file was downloaded/);
    });

    it("fails instead of completing after OWSSVR HTTP or parse errors", () => {
      assert.doesNotMatch(owssvrSource, /resp\.status === 503[\s\S]{0,80}\bbreak;/);
      assert.match(
        owssvrSource,
        /if \(!resp\.ok\) \{[\s\S]{0,500}reportDone\(false, owssvrIncompleteMessage\([\s\S]{0,300}\breturn;/
      );
      assert.match(
        owssvrSource,
        /catch \(e\) \{[\s\S]{0,250}reportDone\(false, owssvrIncompleteMessage\([\s\S]{0,100}\breturn;/
      );
    });

    it("fails instead of treating an ID-page HTTP error as end-of-data", () => {
      assert.doesNotMatch(restSource, /if \(!idResp\.ok\) break;/);
      assert.match(
        restSource,
        /if \(!idResp\.ok\) \{[\s\S]{0,300}reportDone\(false, "Export incomplete while enumerating item IDs[\s\S]{0,200}\breturn;/
      );
    });
  });
}
