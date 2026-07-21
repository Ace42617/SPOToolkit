import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { describe, it } from "node:test";
import assert from "node:assert/strict";

const popupScripts = [
  "../popup.js",
  "../SP-Developer-Toolkit-Lite/popup.js",
  "../SP-Developer-Toolkit-Experimental/popup.js",
];

describe("popup scripts", () => {
  for (const relativePath of popupScripts) {
    it(`${relativePath} parses`, () => {
      const scriptPath = fileURLToPath(new URL(relativePath, import.meta.url));
      const result = spawnSync(process.execPath, ["--check", scriptPath], {
        encoding: "utf8",
      });

      assert.equal(result.status, 0, result.stderr || result.stdout);
    });
  }
});
