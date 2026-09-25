// The file details panel and the slide chrome: what the panel lists for each
// format, that text is not clipped in a slide, and that thumbnails scale to
// their frame.
import { createRequire } from "node:module";
import { readFileSync, existsSync } from "node:fs";
import { setupDom, createContainer } from "./harness.mjs";
import { testPath, skipNotice } from "./env.mjs";

setupDom();
const require = createRequire(import.meta.url);
const { parsePptx } = require("../src/pptx/parse.js");
const { createPptxRenderer } = require("../src/pptx/render.js");
const { parseDocx } = require("../src/docx/parse.js");
const { readWorkbook } = require("../src/spreadsheet/read.js");
const { Package } = require("../src/shared/package.js");

let pass = 0;
let fail = 0;
function check(name, condition, detail) {
  if (condition) { pass++; console.log("ok   " + name); }
  else { fail++; console.log("FAIL " + name + (detail !== undefined ? "  -> " + detail : "")); }
}

const pptxFile = testPath("OV_TEST_PPTX");
if (pptxFile) {
  console.log("--- presentation ---");
  const model = parsePptx(new Uint8Array(readFileSync(pptxFile)));
  const host = createContainer();
  const renderer = createPptxRenderer({ container: host, model, settings: { fit: "none" } });
  renderer.goToSlide(0);

  // Thumbnails scale to their frame instead of a fixed factor.
  const thumbs = host.querySelectorAll(".ov-pptx-thumbbox .ov-pptx-page");
  check("thumbnails are built", thumbs.length > 0, String(thumbs.length));
  if (thumbs.length) {
    const transform = thumbs[0].style.transform || "";
    check("a thumbnail scales to its frame", /scale\(0\.0/.test(transform) || /scale\(0\.1/.test(transform), transform);
  }

  // Text may spill rather than being cut in half.
  const boxes = Array.from(host.querySelectorAll(".ov-pptx-text"));
  check("text boxes are built", boxes.length > 0, String(boxes.length));
  check("a text box does not hide its overflow", boxes.every((el) => el.style.overflow === "visible"), "some clip");

  // Fonts resolve through a fallback stack so a missing face does not change
  // the line breaks.
  const run = host.querySelector(".ov-pptx-r");
  const family = run ? run.style.fontFamily : "";
  check("runs carry a font stack", family.includes("sans-serif") || family.includes("serif"), family);
  renderer.destroy();
} else {
  skipNotice("OV_TEST_PPTX", "presentation checks");
}

if (pptxFile) {
  const model = parsePptx(new Uint8Array(readFileSync(pptxFile)));
  console.log("--- presentation details ---");
  const meta = require("../src/app/presentation-controller.js");
  const host = createContainer();
  const plugin = { settings: {}, updateSetting: () => {} };
  const controller = new meta.PresentationController({ host, plugin, format: "pptx", callbacks: {} });
  controller.load(new Uint8Array(readFileSync(pptxFile)));
  controller.mount();
  const rows = controller.describe();
  const labels = rows.map((row) => row.label);
  check("details list the slides", labels.includes("Slides"), labels.join(","));
  check("details list the slide size", labels.includes("Slide size"), labels.join(","));
  check("details count shapes", labels.includes("Shapes"), labels.join(","));
  check("details report pictures or tables", labels.includes("Pictures") || labels.includes("Tables"), labels.join(","));
  controller.destroy();
}

const docxFile = testPath("OV_TEST_DOCX");
if (docxFile) {
  console.log("--- document details ---");
  const meta = require("../src/app/document-controller.js");
  const host = createContainer();
  const plugin = { settings: {}, updateSetting: () => {} };
  const controller = new meta.DocumentController({ host, plugin, format: "docx", callbacks: {} });
  controller.load(new Uint8Array(readFileSync(docxFile)), "x.docx");
  controller.mount();
  const labels = controller.describe().map((row) => row.label);
  check("details list words", labels.includes("Words"), labels.join(","));
  check("details list characters", labels.includes("Characters"), labels.join(","));
  check("details list rendered pages", labels.includes("Pages (rendered)"), labels.join(","));
  check("details list the page size", labels.includes("Page size"), labels.join(","));
  controller.destroy();
}

const xlsxFile = testPath("OV_TEST_XLSX");
if (xlsxFile) {
  console.log("--- workbook details ---");
  const book = readWorkbook(new Uint8Array(readFileSync(xlsxFile)));
  const meta = require("../src/app/spreadsheet-controller.js");
  const host = createContainer();
  const tabs = globalThis.document.createElement("div");
  const bars = globalThis.document.createElement("div");
  const plugin = { settings: { zoom: 100, sheetBackground: "white", showGridlines: true, showHeaders: true, maxRows: 0 }, updateSetting: () => {}, saveFileState: () => {} };
  const controller = new meta.SpreadsheetController({ host, tabsEl: tabs, filterBarEl: bars, plugin, format: "xlsx", callbacks: {} });
  controller.load(new Uint8Array(readFileSync(xlsxFile)), "book.xlsx");
  controller.mount();
  const labels = controller.describe().map((row) => row.label);
  check("details list the sheets", labels.includes("Sheets"), labels.join(","));
  check("details list the active sheet size", labels.includes("Active sheet size"), labels.join(","));
  const properties = book.properties || {};
  check("the workbook properties were read", Object.keys(properties).length > 0, JSON.stringify(properties).slice(0, 60));
  controller.destroy();
}

// ---------- package metadata ----------

// A synthetic package: only the parts the metadata reader looks for. Nothing
// here comes from a real document.
const { zipSync } = require("fflate");
const { describePackage, sectionsToText } = require("../src/shared/metadata.js");
const xml = (body) => new TextEncoder().encode('<?xml version="1.0"?>' + body);
const synthetic = zipSync({
  "docProps/core.xml": xml('<cp:coreProperties xmlns:cp="x" xmlns:dc="y" xmlns:dcterms="z">' +
    "<dc:title>Quarterly report</dc:title><dc:subject>Finance</dc:subject><dc:creator>Ada Lovelace</dc:creator>" +
    "<cp:lastModifiedBy>Grace Hopper</cp:lastModifiedBy><cp:revision>7</cp:revision><dc:language>en-GB</dc:language>" +
    "<dcterms:created>2024-01-02T03:04:05Z</dcterms:created></cp:coreProperties>"),
  "docProps/app.xml": xml("<Properties><Application>Microsoft Word</Application><AppVersion>16.0</AppVersion>" +
    "<Company>Analytical Engines</Company><Words>1234</Words><TotalTime>42</TotalTime></Properties>"),
  "docProps/custom.xml": xml('<Properties><property name="Project code" pid="2"><vt:lpwstr>AE-1</vt:lpwstr></property>' +
    '<property name="Reviewed" pid="3"><vt:bool>true</vt:bool></property></Properties>'),
  "word/document.xml": xml("<w:document/>"),
  "word/comments.xml": xml('<w:comments><w:comment w:author="Ada Lovelace" w:initials="AL" w:date="2024-01-02T03:04:05Z"/>' +
    '<w:comment w:author="Ada Lovelace" w:initials="AL"/><w:comment w:author="Charles Babbage" w:initials="CB"/></w:comments>'),
  "word/styles.xml": xml('<w:styles><w:docDefaults><w:rPrDefault><w:rPr><w:lang w:val="en-GB"/></w:rPr></w:rPrDefault></w:docDefaults>' +
    '<w:style w:styleId="Normal"/></w:styles>'),
  "word/settings.xml": xml('<w:settings><w:trackChanges/><w:documentProtection w:edit="readOnly" w:enforcement="1"/></w:settings>'),
  "word/media/image1.png": new Uint8Array(2048),
});
const syntheticPkg = Package.open(synthetic);
// meta.mjs's check takes a condition; these compare two values.
const expect = (name, got, want) => check(name, got === want, JSON.stringify(got) + " wanted " + JSON.stringify(want));
const sections = describePackage(syntheticPkg, { kind: "document", model: null });
const byTitle = new Map(sections.map((entry) => [entry.title, entry]));
const rowValue = (title, label) => {
  const entry = byTitle.get(title);
  if (!entry) return null;
  const row = entry.rows.find((item) => item.label === label);
  return row ? row.value : null;
};

expect("metadata: author from core.xml", rowValue("Document properties", "Author"), "Ada Lovelace");
expect("metadata: title", rowValue("Document properties", "Title"), "Quarterly report");
expect("metadata: last modified by", rowValue("Document properties", "Last modified by"), "Grace Hopper");
expect("metadata: language", rowValue("Document properties", "Language"), "en-GB");
expect("metadata: application", rowValue("Application", "Application"), "Microsoft Word");
expect("metadata: company", rowValue("Application", "Company"), "Analytical Engines");
expect("metadata: custom property", rowValue("Custom properties", "Project code"), "AE-1");
expect("metadata: custom boolean", rowValue("Custom properties", "Reviewed"), "true");
expect("metadata: comment authors", rowValue("Comments", "Comment author"), "Ada Lovelace");
expect("metadata: comment count", rowValue("Comments", "Comments"), "3");
expect("metadata: protection", rowValue("Word document", "Protection"), "readOnly, enforced");
expect("metadata: track changes", rowValue("Word document", "Track changes"), "on");
expect("metadata: styles counted", rowValue("Word document", "Styles"), "1");
expect("metadata: package parts", rowValue("Package", "Parts"), String(syntheticPkg.list().length));
expect("metadata: media counted", rowValue("Package", "Media parts"), "1");
check("metadata: as text", sectionsToText(sections).indexOf("Author: Ada Lovelace") !== -1);
const bare = describePackage(Package.open(zipSync({ "a.xml": xml("<a/>") })), {});
check("metadata: no property sections without properties", bare.every((entry) => entry.title !== "Document properties" && entry.title !== "Custom properties"));
check("metadata: the package section still reports parts", bare.some((entry) => entry.title === "Package"));

console.log("");
console.log("meta:", pass, "pass,", fail, "fail");
process.exit(fail ? 1 : 0);
