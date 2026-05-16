import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { existsSync } from "node:fs";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const liteRoot = path.join(repoRoot, "SP-Developer-Toolkit-Lite");

function assertInsideLitePackage(resolvedPath, specifier) {
  const relative = path.relative(liteRoot, resolvedPath);
  assert.ok(
    relative && !relative.startsWith("..") && !path.isAbsolute(relative),
    `Lite popup import ${specifier} resolves outside the Lite package`
  );
}

describe("Lite package module imports", () => {
  it("keeps popup static imports inside the unpacked Lite extension root", async () => {
    const popupPath = path.join(liteRoot, "popup.js");
    const popupSource = await readFile(popupPath, "utf8");
    const importPattern = /\bimport\s+(?:[^'"]+\s+from\s+)?["']([^"']+)["']/g;
    const imports = [...popupSource.matchAll(importPattern)].map((match) => match[1]);

    assert.ok(imports.length > 0, "expected popup.js to have static imports");
    for (const specifier of imports) {
      if (!specifier.startsWith(".")) continue;
      assert.ok(!specifier.startsWith("../"), `Lite popup import ${specifier} escapes the extension root`);
      const resolved = path.resolve(path.dirname(popupPath), specifier);
      assertInsideLitePackage(resolved, specifier);
      assert.ok(existsSync(resolved), `Lite popup import ${specifier} is not packaged`);
    }
  });
});
