import fs from "fs";
import { createRequire } from "module";
const require = createRequire(import.meta.url);
const JSZip = require("../jszip.min.js");

async function stats(p) {
  const buf = fs.readFileSync(p);
  const zip = await JSZip.loadAsync(buf);
  const wb = await zip.file("xl/workbook.xml").async("string");
  const names = [...wb.matchAll(/name="([^"]+)"/g)].map((m) => m[1]);
  const out = { file: p.split(/[/\\]/).pop(), sheets: {} };
  for (let i = 0; i < names.length; i++) {
    const xml = await zip.file(`xl/worksheets/sheet${i + 1}.xml`).async("string");
    out.sheets[names[i]] = (xml.match(/<row /g) || []).length;
  }
  return out;
}

const files = [
  "c:/Users/Alex/Downloads/Calista Brice-PermissionsMatrix-20260604-195731.xlsx",
  "c:/Temp/Calista Brice-PermissionsMatrix-20260603-224121.xlsx"
];
for (const f of files) {
  console.log(JSON.stringify(await stats(f), null, 2));
}
