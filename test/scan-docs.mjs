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
 * ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 * FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
 * details.
 *
 * You should have received a copy of the GNU General Public License along with
 * this program. If not, see <https://www.gnu.org/licenses/>.
 */

// What is actually inside a folder of Office documents: the metadata editors
// leave behind, and the content features that decide what a viewer still has to
// learn. Only the ZIP directory and the XML parts are read, so a 200 MB deck
// costs a few milliseconds.
//
//   bun test/scan-docs.mjs [--folder DIR] [--json OUT] [--limit N] [paths...]
//
// Like the other file backed tools it takes its default folder from .testenv
// (OV_TEST_FOLDER), so no personal path is ever committed. The output names only
// basenames.

import { createRequire } from "node:module";
import { readdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { basename, join } from "node:path";
import { setupDom } from "./harness.mjs";
import { testPath, skipNotice } from "./env.mjs";

setupDom();
const require = createRequire(import.meta.url);
const { Package } = require("../src/shared/package.js");

const SUPPORTED = /\.(xlsx|xlsm|xltx|xltm|csv|tsv|ods|docx|docm|dotx|dotm|odt|rtf|pptx|pptm|ppsx|ppsm|potx|potm|odp)$/i;
const OLE = /\.(doc|xls|ppt|pps|wps)$/i;

const args = process.argv.slice(2);
function flagValue(name) {
  const i = args.indexOf(name);
  return i === -1 ? null : args[i + 1];
}
const limit = Number(flagValue("--limit")) || Infinity;
const jsonOut = flagValue("--json");
let folder = flagValue("--folder");
// Positional paths only: the value of a flag is not a path.
const flagValues = new Set([folder, jsonOut, flagValue("--limit")].filter(Boolean));
const explicit = [];
for (let i = 0; i < args.length; i++) {
  const arg = args[i];
  if (arg.startsWith("--")) {
    i++;
    continue;
  }
  if (!flagValues.has(arg)) explicit.push(arg);
}
if (!folder && !explicit.length) folder = testPath("OV_TEST_FOLDER");
if (!folder && !explicit.length) {
  skipNotice("OV_TEST_FOLDER", "feature scan");
  process.exit(0);
}

function walk(dir, out, depth) {
  let entries;
  try {
    entries = readdirSync(dir, { withFileTypes: true });
  } catch (err) {
    return out;
  }
  if ((depth || 0) > 12) return out;
  for (const entry of entries) {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) {
      if (entry.name === "node_modules" || entry.name.startsWith(".")) continue;
      walk(path, out, (depth || 0) + 1);
    } else if (entry.isFile() && (SUPPORTED.test(entry.name) || OLE.test(entry.name))) {
      out.push(path);
    }
  }
  return out;
}

// Every feature the scan knows how to spot, and what the viewer does with it.
// kind: full, partial, placeholder, none.
const FEATURES = {
  // shared
  coreProps: { label: "Core properties (author, title, revision)", kind: "full" },
  appProps: { label: "Application properties", kind: "full" },
  customProps: { label: "Custom properties", kind: "full" },
  language: { label: "Document language", kind: "full" },
  thumbnail: { label: "Embedded thumbnail", kind: "partial", note: "shown in the details panel only" },
  macros: { label: "Macros (vbaProject)", kind: "partial", note: "ignored, file still opens" },
  signatures: { label: "Digital signatures", kind: "partial", note: "detected, not verified" },
  externalRels: { label: "External relationships", kind: "full", note: "links open in the system browser" },
  embeddedOle: { label: "Embedded OLE objects", kind: "partial", note: "frame kept, content not drawn" },
  embeddedPackage: { label: "Embedded files", kind: "none" },
  comments: { label: "Comments", kind: "partial", note: "counted with authors, not drawn" },
  threadedComments: { label: "Threaded comments", kind: "partial", note: "counted with authors" },
  ink: { label: "Ink annotations", kind: "none" },
  threeD: { label: "3D models", kind: "none" },
  chartEx: { label: "Modern charts (waterfall, funnel, treemap)", kind: "none" },
  // word
  footnotes: { label: "Footnotes and endnotes", kind: "full" },
  headersFooters: { label: "Headers and footers", kind: "full" },
  trackedChanges: { label: "Tracked changes", kind: "partial", note: "detected, revisions shown as content" },
  protection: { label: "Document protection", kind: "partial", note: "detected, shown read only anyway" },
  equations: { label: "Equations (OMML)", kind: "partial", note: "linear text, not typeset" },
  vmlPictures: { label: "Legacy VML pictures", kind: "full" },
  textBoxes: { label: "Text boxes and shapes", kind: "full" },
  contentControls: { label: "Content controls", kind: "partial", note: "contents drawn, control chrome not" },
  wordCharts: { label: "Word charts", kind: "placeholder" },
  smartArt: { label: "SmartArt / diagrams", kind: "placeholder" },
  fields: { label: "Fields (PAGE, REF, TOC)", kind: "partial", note: "PAGE and NUMPAGES resolve, others keep saved text" },
  // presentation
  slideCharts: { label: "Real chart parts", kind: "placeholder" },
  diagrams: { label: "SmartArt / diagrams", kind: "placeholder" },
  transitions: { label: "Slide transitions", kind: "none", note: "static viewer" },
  animations: { label: "Animations", kind: "none", note: "static viewer" },
  audioVideo: { label: "Audio and video", kind: "partial", note: "icon drawn, playback not offered" },
  embeddedFonts: { label: "Embedded fonts", kind: "none", note: "fallback metrics used" },
  notes: { label: "Speaker notes", kind: "full" },
  slideSections: { label: "Slide sections", kind: "none" },
  // spreadsheet
  xlsxCharts: { label: "Workbook charts", kind: "placeholder" },
  pivotTables: { label: "Pivot tables", kind: "none", note: "the underlying sheet still shows" },
  conditionalFormatting: { label: "Conditional formatting", kind: "full" },
  dataValidation: { label: "Data validation", kind: "none" },
  sparklines: { label: "Sparklines", kind: "none" },
  slicers: { label: "Slicers and timelines", kind: "none" },
  richTextCells: { label: "Rich text in cells", kind: "partial", note: "cell text drawn, runs flattened" },
  arrayFormulas: { label: "Array and shared formulas", kind: "partial", note: "cached values shown" },
  connections: { label: "Data connections and queries", kind: "none" },
  tables: { label: "Tables (list objects)", kind: "partial", note: "banding from styles, filter buttons drawn" },
  // odf / legacy
  legacyBinary: { label: "Legacy binary formats (.doc, .xls, .ppt)", kind: "none", note: "refused with an explanation" },
};

// A field counts when it carries a value: an empty <dc:title></dc:title> is not
// metadata anyone wants to see.
function hasField(text, name) {
  return new RegExp("<(?:[A-Za-z0-9]+:)?" + name + "[^>]*>\\s*[^<\\s]", "i").test(text);
}

function bump(map, key, file) {
  if (!FEATURES[key]) return;
  if (!map.has(key)) map.set(key, { count: 0, files: [] });
  const entry = map.get(key);
  entry.count++;
  if (entry.files.length < 4 && entry.files.indexOf(file) === -1) entry.files.push(file);
}

const found = new Map();
const mediaTypes = new Map();
const languages = new Map();
const coreFields = new Map();
const appFields = new Map();
const errors = [];
const files = [];
for (const path of explicit) files.push(path);
if (folder) walk(folder, files, 0);
const seen = [];
for (const path of files) {
  if (seen.length >= limit) break;
  seen.push(path);
}

for (const path of seen) {
  const name = basename(path);
  let bytes;
  try {
    bytes = new Uint8Array(readFileSync(path));
  } catch (err) {
    errors.push([name, "unreadable"]);
    continue;
  }
  if (OLE.test(name)) {
    bump(found, "legacyBinary", name);
    continue;
  }
  if (!SUPPORTED.test(name)) continue;
  let pkg;
  try {
    pkg = Package.open(bytes);
  } catch (err) {
    errors.push([name, err.message]);
    continue;
  }
  const parts = pkg.list();
  const has = (re) => parts.some((part) => re.test(part));
  const read = (re) => parts.filter((part) => re.test(part)).map((part) => pkg.text(part) || "").join("\n");

  // shared
  if (pkg.has("docProps/core.xml")) {
    bump(found, "coreProps", name);
    const text = pkg.text("docProps/core.xml") || "";
    for (const field of ["title", "subject", "creator", "lastModifiedBy", "keywords", "description", "category", "revision", "created", "modified"]) {
      if (hasField(text, field)) coreFields.set(field, (coreFields.get(field) || 0) + 1);
    }
    const lang = /<dc:language[^>]*>([^<]*)</.exec(text);
    if (lang) languages.set(lang[1], (languages.get(lang[1]) || 0) + 1);
  }
  if (pkg.has("docProps/app.xml")) {
    bump(found, "appProps", name);
    const text = pkg.text("docProps/app.xml") || "";
    for (const field of ["Application", "AppVersion", "Company", "Manager", "Template", "TotalTime", "Pages", "Words", "Slides", "Notes"]) {
      if (hasField(text, field)) appFields.set(field, (appFields.get(field) || 0) + 1);
    }
  }
  if (pkg.has("docProps/custom.xml")) bump(found, "customProps", name);
  if (has(/^docProps\/thumbnail\./)) bump(found, "thumbnail", name);
  if (has(/vbaProject\.bin$/)) bump(found, "macros", name);
  if (has(/^_xmlsignatures\//)) bump(found, "signatures", name);
  if (has(/^word\/embeddings\//) || has(/^xl\/embeddings\//) || has(/^ppt\/embeddings\//)) bump(found, "embeddedOle", name);
  if (has(/\/media\/.*\.(bin|dat)$/i)) bump(found, "embeddedPackage", name);
  if (has(/^word\/comments\.xml$/i) || has(/^xl\/comments\d*\.xml$/i) || has(/^ppt\/comments\//)) bump(found, "comments", name);
  if (has(/threadedComment/i)) bump(found, "threadedComments", name);
  if (has(/\/ink\//i) || has(/ink\.xml$/i)) bump(found, "ink", name);
  if (has(/model\/gltf|\.glb$/i)) bump(found, "threeD", name);
  if (has(/chartEx|chartex/i)) bump(found, "chartEx", name);
  // media inventory
  for (const part of parts) {
    if (!/\/media\//.test(part)) continue;
    const ext = (part.split(".").pop() || "").toLowerCase();
    mediaTypes.set(ext, (mediaTypes.get(ext) || 0) + 1);
  }
  // external relationships
  if (parts.some((part) => /\.rels$/.test(part) && /TargetMode="External"/.test(pkg.text(part) || ""))) bump(found, "externalRels", name);

  if (/^word\//.test(parts[0] || "") || parts.some((p) => p.startsWith("word/"))) {
    const documentText = pkg.text("word/document.xml") || "";
    const settings = pkg.text("word/settings.xml") || "";
    if (/<w:footnoteReference\b|<w:endnoteReference\b/.test(documentText) || pkg.has("word/footnotes.xml")) bump(found, "footnotes", name);
    if (has(/^word\/header\d*\.xml$/) || has(/^word\/footer\d*\.xml$/)) bump(found, "headersFooters", name);
    if (/<w:(ins|del)\b/.test(documentText)) bump(found, "trackedChanges", name);
    if (/<w:documentProtection\b/.test(settings)) bump(found, "protection", name);
    if (/<m:oMath\b|<m:oMathPara\b/.test(documentText)) bump(found, "equations", name);
    if (/<w:pict\b/.test(documentText)) bump(found, "vmlPictures", name);
    if (/<wps:txbx\b|<v:textbox\b/.test(documentText)) bump(found, "textBoxes", name);
    if (/<w:sdt\b/.test(documentText)) bump(found, "contentControls", name);
    if (has(/^word\/charts\//)) bump(found, "wordCharts", name);
    if (has(/^word\/diagrams\//)) bump(found, "smartArt", name);
    if (/<w:fldChar\b|<w:instrText\b/.test(documentText)) bump(found, "fields", name);
  }

  if (parts.some((p) => p.startsWith("ppt/"))) {
    if (has(/^ppt\/charts\//)) bump(found, "slideCharts", name);
    if (has(/^ppt\/diagrams\//)) bump(found, "diagrams", name);
    if (has(/^ppt\/notesSlides\//)) bump(found, "notes", name);
    if (has(/\.fntdata$/)) bump(found, "embeddedFonts", name);
    const slideText = read(/^ppt\/slides\/slide\d+\.xml$/);
    if (/<p:transition\b/.test(slideText)) bump(found, "transitions", name);
    if (/<p:timing\b/.test(slideText)) bump(found, "animations", name);
    if (/<p14:section\b|<p:section\b/.test(pkg.text("ppt/presentation.xml") || "")) bump(found, "slideSections", name);
    if (has(/^(ppt|word|xl)\/media\/.*\.(mp3|m4a|wav|mp4|mov|avi|wmv|webm)$/i)) bump(found, "audioVideo", name);
  }

  if (parts.some((p) => p.startsWith("xl/"))) {
    if (has(/^xl\/charts\//)) bump(found, "xlsxCharts", name);
    if (has(/^xl\/pivotTables\//)) bump(found, "pivotTables", name);
    if (has(/^xl\/tables\//)) bump(found, "tables", name);
    if (has(/^xl\/connections\.xml$/) || has(/^xl\/queryTables\//)) bump(found, "connections", name);
    if (has(/^xl\/slicers?\//) || has(/^xl\/timelines?\//)) bump(found, "slicers", name);
    const sheetText = read(/^xl\/worksheets\/sheet\d+\.xml$/);
    if (/<conditionalFormatting\b/.test(sheetText)) bump(found, "conditionalFormatting", name);
    if (/<dataValidations?\b/.test(sheetText)) bump(found, "dataValidation", name);
    if (/<x14:sparklineGroups?\b|<x14:sparklineGroup\b/.test(sheetText)) bump(found, "sparklines", name);
    if (/<f[^>]*t="array"/.test(sheetText)) bump(found, "arrayFormulas", name);
    const strings = pkg.text("xl/sharedStrings.xml") || "";
    if (/<si>\s*<r>/.test(strings)) bump(found, "richTextCells", name);
  }
}

// The report: prevalence first, then the gap list.
const rows = [];
for (const [key, entry] of found) rows.push({ key, ...FEATURES[key], count: entry.count, files: entry.files });
rows.sort((a, b) => b.count - a.count || a.key.localeCompare(b.key));

console.log(`scanned ${seen.length} files${folder ? " under " + basename(folder) : ""}`);
if (errors.length) {
  console.log(`unreadable or refused: ${errors.length}`);
  for (const [name, reason] of errors.slice(0, 6)) console.log("   " + name + ": " + reason);
}
console.log("");
console.log("feature                        files   viewer");
for (const row of rows) {
  const state = row.kind === "full" ? "yes" : row.kind === "partial" ? "partly" : row.kind === "placeholder" ? "placeholder" : "NO";
  console.log("  " + row.label.padEnd(30) + String(row.count).padStart(5) + "   " + state.padEnd(11) + (row.note ? row.note : ""));
}
const gaps = rows.filter((row) => row.kind === "none" || row.kind === "placeholder");
if (gaps.length) {
  console.log("");
  console.log("gaps, most common first:");
  for (const gap of gaps) console.log("  " + gap.label + " (" + gap.count + ")  e.g. " + gap.files[0]);
}
console.log("");
console.log("core property fields:", Array.from(coreFields).map(([k, v]) => k + " " + v).join(", ") || "none");
console.log("app property fields :", Array.from(appFields).map(([k, v]) => k + " " + v).join(", ") || "none");
console.log("metadata languages  :", Array.from(languages).map(([k, v]) => k + " " + v).join(", ") || "none");
console.log("media types         :", Array.from(mediaTypes).map(([k, v]) => k + " x" + v).join(", ") || "none");

if (jsonOut) {
  writeFileSync(jsonOut, JSON.stringify({
    scanned: seen.length,
    features: rows,
    gaps: gaps.map((gap) => gap.key),
    coreFields: Object.fromEntries(coreFields),
    appFields: Object.fromEntries(appFields),
    languages: Object.fromEntries(languages),
    mediaTypes: Object.fromEntries(mediaTypes),
    errors,
  }, null, 2));
  console.log("wrote " + jsonOut);
}
