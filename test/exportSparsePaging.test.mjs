import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const exportSource = await readFile(new URL("../exportCSV.js", import.meta.url), "utf8");
const helperSource = exportSource.match(
  /function shouldStopAfterEmptyIdRange\([^)]*\) \{[\s\S]*?\n  \}/
);

assert.ok(helperSource, "exportCSV.js should define the empty-range termination helper");

const shouldStopAfterEmptyIdRange = Function(`return (${helperSource[0]});`)();

describe("OWSSVR sparse ID-range paging", () => {
  it("keeps scanning empty ranges until a resolved maximum item ID", () => {
    assert.equal(
      shouldStopAfterEmptyIdRange(50_001, 80_000, 51, 2_000),
      false,
      "a 50,000-ID deletion gap must not truncate items that survive above it"
    );
    assert.equal(shouldStopAfterEmptyIdRange(79_001, 80_000, 80, 2_000), false);
    assert.equal(shouldStopAfterEmptyIdRange(80_001, 80_000, 81, 2_000), true);
  });

  it("retains a bounded fallback when the maximum ID cannot be resolved", () => {
    assert.equal(shouldStopAfterEmptyIdRange(1_999_001, null, 1_999, 2_000), false);
    assert.equal(shouldStopAfterEmptyIdRange(2_000_001, null, 2_000, 2_000), true);
  });

  it("uses the termination helper in the empty-page branch", () => {
    assert.match(
      exportSource,
      /if \(shouldStopAfterEmptyIdRange\(startId, maxIdResolved, emptyStreak, maxEmptyStreak\)\) break;/
    );
  });
});
