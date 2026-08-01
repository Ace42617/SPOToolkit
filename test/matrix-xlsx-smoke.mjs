import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import vm from "vm";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const jszipCode = fs.readFileSync(path.join(__dirname, "..", "jszip.min.js"), "utf8");
const ctx = {
  console,
  Buffer,
  setTimeout,
  clearTimeout,
  setImmediate,
  clearImmediate,
};
vm.runInNewContext(jszipCode + "\nthis.JSZip = JSZip;", ctx);
const JSZip = ctx.JSZip;

function escXml(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
function colToLetter(c) {
  let s = "";
  for (;;) {
    s = String.fromCharCode((c % 26) + 65) + s;
    c = Math.floor(c / 26) - 1;
    if (c < 0) break;
  }
  return s;
}

const sheets = [
  { name: "Summary", aoa: [["A", "B"], ["1", "2"]] },
  { name: "All Items", aoa: [["H1", "H2"], ...Array.from({ length: 80000 }, (_, i) => ["site" + i, "/path/" + i])] }
];
const sharedStrings = [];
const sharedIndex = {};
let totalStringRefs = 0;
function getStrIndex(val) {
  const s = val == null ? "" : String(val);
  totalStringRefs++;
  if (sharedIndex[s] === undefined) {
    sharedIndex[s] = sharedStrings.length;
    sharedStrings.push(s);
  }
  return sharedIndex[s];
}
function buildSheetRows(aoa, sparse) {
  const nr = aoa.length;
  const nc = nr > 0 ? aoa[0].length : 0;
  const sheetRows = [];
  for (let r = 0; r < nr; r++) {
    const row = aoa[r];
    const cells = [];
    for (let c = 0; c < nc; c++) {
      const v = row[c];
      if (sparse && r > 0 && (v == null || v === "")) continue;
      const ref = colToLetter(c) + (r + 1);
      cells.push('<c r="' + ref + '" t="s"><v>' + getStrIndex(v) + "</v></c>");
    }
    if (cells.length) sheetRows.push('<row r="' + (r + 1) + '">' + cells.join("") + "</row>");
  }
  return { rows: sheetRows, nr, nc };
}

const sheetXmls = [];
for (const aoa of sheets.map((s) => s.aoa)) {
  const sparse = aoa.length > 5000;
  const built = buildSheetRows(aoa, sparse);
  const filterRange = "A1:" + colToLetter(built.nc - 1) + built.nr;
  sheetXmls.push(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="' +
      filterRange +
      '"/><sheetData>' +
      built.rows.join("") +
      "</sheetData></worksheet>"
  );
}
const sstItems = sharedStrings.map((s) => "<si><t>" + escXml(s) + "</t></si>");
const sstXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' +
  totalStringRefs +
  '" uniqueCount="' +
  sharedStrings.length +
  '">' +
  sstItems.join("") +
  "</sst>";
const stylesXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>';

const zip = new JSZip();
zip.file(
  "[Content_Types].xml",
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
);
zip.file(
  "_rels/.rels",
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
);
zip.file(
  "xl/workbook.xml",
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Summary" sheetId="1" r:id="rId1"/><sheet name="All Items" sheetId="2" r:id="rId2"/></sheets></workbook>'
);
zip.file(
  "xl/_rels/workbook.xml.rels",
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
);
zip.file("xl/worksheets/sheet1.xml", sheetXmls[0]);
zip.file("xl/worksheets/sheet2.xml", sheetXmls[1]);
zip.file("xl/sharedStrings.xml", sstXml);
zip.file("xl/styles.xml", stylesXml);

const outPath = path.join(process.env.TEMP || "c:\\Temp", "pm-matrix-smoke.xlsx");
const buf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 6 } });
fs.writeFileSync(outPath, buf);
console.log("written", outPath, "bytes", buf.length, "refs", totalStringRefs, "unique", sharedStrings.length);
