/*
 * Styled MS Office Viewer, an Obsidian plugin that renders office documents
 * (xlsx, docx, pptx and their relatives) with their real styling, read only.
 *
 * Copyright (C) 2026 Zoroaster1x
 *
 * This program is free software: you can redistribute it and/or modify it under
 * the terms of the GNU General Public License as published by the Free Software
 * Foundation, either version 3 of the License, or (at your option) any later
 * version.
 *
 * This program is distributed in the hope that it will be useful, but WITHOUT
 * ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
 * FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
 * details.
 *
 * You should have received a copy of the GNU General Public License along with
 * this program. If not, see <https://www.gnu.org/licenses/>.
 */

// OpenDocument routing and parsing. Every ODF file is a zip holding
// content.xml, so the signature alone cannot tell a text document from a
// workbook: the extension has to decide the reader. This suite builds the
// smallest package of each class in memory and checks both halves, so a
// .odt can never be handed to the spreadsheet reader again.
//
//   bun test/odf.mjs

import { createRequire } from "node:module";
import { zipSync, strToU8 } from "fflate";
import { setupDom } from "./harness.mjs";
const require = createRequire(import.meta.url);
setupDom();
const { detectFormat } = require("../src/app/format-router.js");
const { readOds, readOdt, readOdp } = require("../src/odf/parse.js");

let pass = 0;
let fail = 0;
function check(name, got, want) {
  const same = JSON.stringify(got) === JSON.stringify(want);
  if (same) {
    pass++;
    return;
  }
  fail++;
  console.log("FAIL", name, "->", JSON.stringify(got), "wanted", JSON.stringify(want));
}

const NS = [
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"',
  'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"',
  'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
  'xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"',
].join(" ");

function packageOf(mimetype, body, automaticStyles) {
  const content = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content ${NS} office:version="1.3">
<office:automatic-styles>${automaticStyles || ""}</office:automatic-styles>
<office:body>${body}</office:body>
</office:document-content>`;
  const manifest = `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">
<manifest:file-entry manifest:full-path="/" manifest:media-type="${mimetype}"/>
<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
</manifest:manifest>`;
  const styles = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles ${NS} office:version="1.3"><office:styles/>
<office:automatic-styles/></office:document-styles>`;
  return new Uint8Array(zipSync({
    mimetype: strToU8(mimetype),
    "META-INF/manifest.xml": strToU8(manifest),
    "styles.xml": strToU8(styles),
    "content.xml": strToU8(content),
  }));
}

// ---------- a text document ----------

const ODT_MIME = "application/vnd.oasis.opendocument.text";
const odt = packageOf(ODT_MIME, `<office:text>
<text:p>Hello from a text document.</text:p>
<text:p text:style-name="Title">A title with its own style</text:p>
<text:p>Mixed <text:span text:style-name="Em">content</text:span> here.</text:p>
</office:text>`, `<style:style style:name="Title" style:family="paragraph">
<style:paragraph-properties fo:text-align="center"/>
<style:text-properties fo:font-weight="bold" fo:font-size="28pt"/>
</style:style>
<style:style style:name="Em" style:family="text">
<style:text-properties fo:font-style="italic"/>
</style:style>`);

check("odt routes to the document reader", detectFormat("notes.odt", odt).kind, "document");
check("odt keeps its extension", detectFormat("notes.odt", odt).ext, "odt");

// ---------- a spreadsheet, with the row properties that used to throw ----------

const ODS_MIME = "application/vnd.oasis.opendocument.spreadsheet";
const ods = packageOf(ODS_MIME, `<office:spreadsheet>
<table:table table:name="Timetable">
<table:table-column table:style-name="co1"/>
<table:table-row table:style-name="ro1">
<table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell>
<table:table-cell office:value-type="float" office:value="2"><text:p>2</text:p></table:table-cell>
</table:table-row>
</table:table>
</office:spreadsheet>`, `<style:style style:name="ro1" style:family="table-row">
<style:table-row-properties style:row-height="0.75in"/>
</style:style>
<style:style style:name="co1" style:family="table-column">
<style:table-column-properties style:column-width="1.2in"/>
</style:style>`);

check("ods routes to the spreadsheet reader", detectFormat("timetable.ods", ods).kind, "spreadsheet");

let sheet = null;
let odsError = null;
try {
  const book = readOds(ods);
  sheet = book.loadSheet(0);
} catch (err) {
  odsError = err && err.message ? err.message : String(err);
}
check("ods parses a sheet", Boolean(sheet), true);
check("ods names the sheet", sheet && sheet.name, "Timetable");
check("ods keeps two cells", sheet && sheet.cells.size, 2);
check("ods keeps the cell value", sheet && sheet.cells.get("1:1").v, 1);
// The bug behind issue #2: the row style was read from a map that holds parsed
// style properties, while ODF keeps the row element itself.
check("ods reads the row height from its style", sheet && Math.round(sheet.rows.get(1).ht), 54);
check("ods reads the column width from its style", sheet && Math.round(sheet.cols[0].widthPx), 115);

// ---------- a presentation ----------

const ODP_MIME = "application/vnd.oasis.opendocument.presentation";
const odp = packageOf(ODP_MIME, `<office:presentation>
<draw:page draw:name="Slide 1">
<draw:frame svg:width="4in" svg:height="1in" svg:x="0.5in" svg:y="0.5in">
<draw:text-box><text:p>Slide text</text:p></draw:text-box>
</draw:frame>
</draw:page>
</office:presentation>`);

check("odp routes to the presentation reader", detectFormat("talk.odp", odp).kind, "presentation");

let slides = null;
let odpError = null;
try {
  slides = readOdp(odp).slides.length;
} catch (err) {
  odpError = err && err.message ? err.message : String(err);
}
check("odp parses its slides", slides, 1);

// ---------- the same packages with the wrong name ----------

// A mislabelled file still lands on the right reader: the zip signature and
// the parts inside decide, with the declared kind only as a fallback.
check("a workbook named .docx is still a workbook", detectFormat("book.docx", ods).kind, "spreadsheet");
check("an ODF package with no usable extension is a workbook", detectFormat("thing.zip", ods).kind, "spreadsheet");

// ---------- the text reader still reads ----------

let doc = null;
let odtError = null;
try {
  doc = readOdt(odt);
} catch (err) {
  odtError = err && err.message ? err.message : String(err);
}
check("odt parses its paragraphs", Boolean(doc), true);
// Plain text in ODF is a text node inside the paragraph, not an element: a
// reader that walks elements alone produces empty paragraphs and a blank page.
check("odt keeps the paragraph text", (doc.body[0].runs || []).map((run) => run.text).join(""), "Hello from a text document.");
const titleBlock = doc.body[1];
check("odt carries the paragraph style onto its runs", Boolean(titleBlock.runs[0].props && titleBlock.runs[0].props.bold), true);
check("odt keeps the styled size", titleBlock.runs[0].props && titleBlock.runs[0].props.sizeHalfPt, 56);
const mixed = (doc.body[2].runs || []).map((run) => run.text).join("");
check("odt keeps text around a span", mixed, "Mixed content here.");
check("odt keeps the span style", Boolean(doc.body[2].runs[1] && doc.body[2].runs[1].props && doc.body[2].runs[1].props.italic), true);

console.log(`odf: ${pass} pass, ${fail} fail`);
if (fail) process.exit(1);
