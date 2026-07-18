import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { describe, it } from "node:test";
import assert from "node:assert/strict";

const popupScripts = [
  "../popup.js",
  "../SP-Developer-Toolkit-Lite/popup.js",
  "../SP-Developer-Toolkit-Experimental/popup.js",
];

describe("extension popup scripts", () => {
  for (const relativePath of popupScripts) {
    it(`${relativePath} parses as JavaScript`, () => {
      const scriptPath = fileURLToPath(new URL(relativePath, import.meta.url));
      const result = spawnSync(process.execPath, ["--check", scriptPath], {
        encoding: "utf8",
      });

      assert.equal(
        result.status,
        0,
        `Syntax check failed:\n${result.stderr || result.stdout}`
      );
    });
  }
});
