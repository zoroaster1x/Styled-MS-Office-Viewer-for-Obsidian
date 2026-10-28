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

// Converts a folder of documents to OpenDocument with LibreOffice, then checks
// that every word of each converted file appears on the page this viewer draws.
// A blank page, a dropped table or a paragraph that never reached the renderer
// all show up as missing words.
//
//   bun test/odt-sweep.mjs ~/Documents /tmp/odt-sweep --count 24
//
// LibreOffice is required for the conversion; without it the sweep skips.

import { createRequire } from "node:module";
import { readFileSync, readdirSync, statSync, mkdirSync, existsSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { basename, extname, join } from "node:path";
import { setupDom, createContainer } from "./harness.mjs";

setupDom();
const require = createRequire(import.meta.url);
const { readOdt } = require("../src/odf/parse.js");
const { DocumentController } = require("../src/app/document-controller.js");

const folder = process.argv[2];
const outDir = process.argv[3] || "/tmp/odt-sweep";
const countArg = process.argv.indexOf("--count");
const count = countArg === -1 ? 24 : Number(process.argv[countArg + 1]);
if (!folder) {
  console.log("usage: bun test/odt-sweep.mjs <folder> [outdir] [--count N]");
  process.exit(2);
}
mkdirSync(outDir, { recursive: true });

const soffice = spawnSync("sh", ["-c", "command -v soffice || command -v libreoffice"], { encoding: "utf8" });
const binary = (soffice.stdout || "").trim();
if (!binary) {
  console.log("skip: LibreOffice is not installed");
  process.exit(0);
}

// ---------- pick documents to convert ----------

const CONVERTIBLE = new Set(["docx", "doc", "rtf", "html", "htm", "txt", "pptx", "xlsx"]);
function walk(dir, out) {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    if (entry.name.startsWith(".")) continue;
    const full = join(dir, entry.name);
    if (entry.isDirectory()) {
      walk(full, out);
      continue;
    }
    const ext = extname(entry.name).slice(1).toLowerCase();
    if (CONVERTIBLE.has(ext)) out.push({ path: full, ext });
  }
}

const all = [];
walk(folder, all);
// Spread the sample over the tree instead of taking one folder's worth.
const step = Math.max(1, Math.floor(all.length / count));
const sample = [];
for (let i = 0; i < all.length && sample.length < count; i += step) sample.push(all[i]);

console.log("converting", sample.length, "documents with LibreOffice");
const converted = [];
for (const file of sample) {
  const run = spawnSync(binary, ["--headless", "--convert-to", "odt", "--outdir", outDir, file.path], { encoding: "utf8", timeout: 300000 });
  const out = join(outDir, basename(file.path, extname(file.path)) + ".odt");
  if (!existsSync(out)) {
    console.log("  convert failed:", basename(file.path), String(run.stderr || "").trim().slice(0, 80));
    continue;
  }
  converted.push({ source: file.path, odt: out });
}
console.log("converted", converted.length);

// ---------- compare what the file says with what the page shows ----------

// textContent glues blocks together ("alcoholtap water"), which undercounts
// words and reads as missing text. Walk the nodes and separate them.
function visibleText(root) {
  const parts = [];
  const walk = (node) => {
    if (node.nodeType === 3) {
      const text = node.nodeValue || "";
      if (text.trim()) parts.push(text);
      return;
    }
    if (node.nodeType !== 1) return;
    const tag = String(node.tagName || "").toLowerCase();
    if (tag === "br") {
      parts.push(" ");
      return;
    }
    for (const child of node.childNodes || []) walk(child);
    parts.push(" ");
  };
  walk(root);
  return parts.join(" ");
}

function wordsOf(text) {
  return String(text || "")
    .replace(/[\u00ad]/g, "")
    .split(/\s+/)
    .map((w) => w.replace(/^[^\p{L}\p{N}]+|[^\p{L}\p{N}]+$/gu, ""))
    .filter((w) => w.length > 2);
}

function mount(odtPath) {
  const bytes = new Uint8Array(readFileSync(odtPath));
  const host = createContainer();
  const tabsEl = createContainer();
  const filterBarEl = createContainer();
  const settings = {
    zoom: 100, slideZoom: 100, sheetBackground: "white", showGridlines: true, showHeaders: true,
    showSpeakerNotes: true, slideFit: "contain", minColumnWidth: 0, minRowHeight: 0, maxRows: 0, fileState: {},
  };
  const plugin = { settings, savePluginData() {}, saveFileState() {}, updateSetting() {}, registerEvent() {} };
  const controller = new DocumentController({ host, plugin, format: "odt", callbacks: {} });
  controller.filePath = odtPath;
  controller.savedState = {};
  controller.load(bytes, basename(odtPath));
  if (typeof controller.restoreState === "function") controller.restoreState({});
  controller.mount();
  return { host, controller };
}

console.log("");
console.log("file".padEnd(46), "words", "shown", "cover", "verdict");
let bad = 0;
for (const { source, odt } of converted) {
  const model = readOdt(new Uint8Array(readFileSync(odt)));
  const expected = [];
  for (const block of model.body) {
    if (block.type === "p") {
      for (const run of block.runs || []) if (run.type === "text") expected.push(run.text);
    } else if (block.type === "table") {
      // Tables carry their own cells; count their words too.
      for (const row of block.rows || []) {
        for (const cell of row.cells || []) {
          for (const para of cell.blocks || cell.paragraphs || []) {
            for (const run of para.runs || []) if (run.type === "text") expected.push(run.text);
          }
        }
      }
    }
  }
  const wanted = wordsOf(expected.join(" "));
  let shownText = "";
  let failure = null;
  let words = wanted.length;
  try {
    const { host, controller } = mount(odt);
    shownText = visibleText(host);
    controller.destroy && controller.destroy();
  } catch (err) {
    failure = (err && err.message) || String(err);
  }
  const shownWords = new Set(wordsOf(shownText));
  const hit = wanted.filter((w) => shownWords.has(w)).length;
  const cover = words ? hit / words : 1;
  const verdict = failure ? "CRASH " + failure.slice(0, 40) : words === 0 ? "no words in file" : cover < 0.95 ? "MISSING WORDS" : "ok";
  if (failure || (words > 0 && cover < 0.95)) bad++;
  console.log(
    basename(source).slice(0, 44).padEnd(46),
    String(words).padStart(5),
    String(hit).padStart(5),
    (cover * 100).toFixed(0).padStart(5) + "%",
    verdict,
  );
  if (verdict === "MISSING WORDS") {
    const missing = wanted.filter((w) => !shownWords.has(w)).slice(0, 8);
    console.log("    missing: " + missing.join(" "));
  }
}
console.log("");
console.log(bad ? bad + " files need a look" : "no missing words");
