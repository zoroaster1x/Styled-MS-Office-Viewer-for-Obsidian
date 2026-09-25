/*
 * Styled MS Office Viewer, an Obsidian plugin that renders office documents (xlsx,
 * docx, pptx and their relatives) with their real styling, read only.
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

// Developer tooling: parses a real document end to end and prints what the
// viewer would draw. Used to verify output against the source files without a
// GUI. Run with:
//
//   bun test/inspect.mjs "/path/to/file.docx" [--html out.html]

import { createRequire } from "node:module";
import { readFileSync, writeFileSync } from "node:fs";
import { setupDom, createContainer } from "./harness.mjs";
import { testPath, skipNotice } from "./env.mjs";

setupDom();
const require = createRequire(import.meta.url);

const args = process.argv.slice(2);
let path = args.find((a) => !a.startsWith("--"));
const htmlOut = (() => {
  const i = args.indexOf("--html");
  return i === -1 ? null : args[i + 1];
})();
if (!path) {
  path = testPath("OV_TEST_DOCX") || testPath("OV_TEST_PPTX") || testPath("OV_TEST_XLSX");
  if (!path) {
    skipNotice("OV_TEST_DOCX", "inspect run");
    process.exit(0);
  }
}

const bytes = new Uint8Array(readFileSync(path));
const lower = path.toLowerCase();

function time(label, fn) {
  const start = performance.now();
  const out = fn();
  const ms = performance.now() - start;
  console.log(label.padEnd(22), ms.toFixed(1) + " ms");
  return out;
}

if (lower.endsWith(".docx") || lower.endsWith(".docm")) {
  const { parseDocx } = require("../src/docx/parse.js");
  const { createDocxRenderer } = require("../src/docx/render.js");

  const model = time("parse", () => parseDocx(bytes));
  const stats = collectStats(model);
  console.log("paragraphs".padEnd(22), stats.paragraphs);
  console.log("tables".padEnd(22), stats.tables);
  console.log("images".padEnd(22), stats.images);
  console.log("lists".padEnd(22), stats.lists);
  console.log("headings".padEnd(22), stats.headings);
  console.log("blocks".padEnd(22), stats.blocks);
  console.log("styles".padEnd(22), model.styles.paragraph.size + " paragraph, " + model.styles.character.size + " character");
  console.log("notes".padEnd(22), model.footnotes.size + " footnotes, " + model.endnotes.size + " endnotes");
  console.log("page".padEnd(22), Math.round(model.section.pageWidthTw / 20) + "pt x " + Math.round(model.section.pageHeightTw / 20) + "pt");
  console.log("first text".padEnd(22), JSON.stringify(stats.firstText.slice(0, 120)));

  const host = createContainer();
  let outline = null;
  const renderer = time("render", () => createDocxRenderer({
    container: host,
    model,
    settings: { zoom: 1 },
    onReady: (info) => {
      outline = info.outline;
    },
  }));
  const html = host.innerHTML;
  console.log("html length".padEnd(22), html.length);
  console.log("outline".padEnd(22), outline ? outline.length : 0);
  if (outline) {
    for (const item of outline.slice(0, 12)) {
      console.log("   ", "  ".repeat(item.level - 1) + item.text.slice(0, 70));
    }
  }
  if (htmlOut) {
    writeFileSync(htmlOut, wrapHtml(html));
    console.log("wrote".padEnd(22), htmlOut);
  }
  renderer.destroy();
} else if (lower.endsWith(".pptx") || lower.endsWith(".pptm")) {
  const { parsePptx } = require("../src/pptx/parse.js");
  const model = time("parse", () => parsePptx(bytes));
  console.log("slides".padEnd(22), model.slides.length);
  console.log("size".padEnd(22), Math.round(model.widthPx) + " x " + Math.round(model.heightPx) + " px");
  for (const slide of model.slides.slice(0, 10)) {
    console.log("slide " + (slide.index + 1), slide.shapes.length + " shapes", slide.text.slice(0, 60).replace(/\s+/g, " "));
  }
} else if (lower.endsWith(".xlsx") || lower.endsWith(".xlsm")) {
  const { readWorkbook } = require("../src/spreadsheet/read.js");
  const book = time("parse", () => readWorkbook(bytes));
  console.log("sheets".padEnd(22), book.sheets.length);
  for (const sheet of book.sheets) {
    const model = book.loadSheet(sheet.index);
    console.log(" " + sheet.name, model.dims ? model.dims.r2 + " rows x " + model.dims.c2 + " cols" : "empty");
  }
} else {
  console.error("unknown extension");
  process.exit(2);
}

function collectStats(model) {
  const stats = { paragraphs: 0, tables: 0, images: 0, lists: 0, headings: 0, blocks: 0, firstText: "" };
  const walk = (blocks) => {
    for (const block of blocks || []) {
      stats.blocks++;
      if (block.type === "p") {
        stats.paragraphs++;
        if (block.numbering) stats.lists++;
        if (block.props && block.props.outlineLevel != null) stats.headings++;
        for (const run of block.runs || []) countRun(run, stats);
        if (!stats.firstText) stats.firstText = plainText(block.runs);
      } else if (block.type === "table") {
        stats.tables++;
        for (const row of block.rows) for (const cell of row.cells) walk(cell.blocks);
      }
    }
  };
  walk(model.body);
  return stats;
}

function countRun(run, stats) {
  if (!run) return;
  if (run.type === "image") stats.images++;
  else if (run.type === "run") for (const inner of run.runs) countRun(inner, stats);
  else if (run.type === "link") for (const inner of run.link.runs) countRun(inner, stats);
}

function plainText(runs) {
  let out = "";
  for (const run of runs || []) {
    if (!run) continue;
    if (run.type === "text") out += run.text;
    else if (run.type === "run") out += plainText(run.runs);
    else if (run.type === "link") out += plainText(run.link.runs);
  }
  return out;
}

function wrapHtml(body) {
  return `<!doctype html>
<html><head><meta charset="utf-8"><title>Office Viewer render</title>
<style>
body { margin: 0; background: #f0f0f0; font-family: Arial, sans-serif; }
.ov-docx { --ov-docx-zoom: 1; }
.ov-docx-scroll { overflow: auto; padding: 20px; }
.ov-docx-page { background: #fff; margin: 0 auto 20px; box-shadow: 0 1px 4px rgba(0,0,0,.25); box-sizing: border-box; color: #000; }
.ov-docx-pagecontent { display: flow-root; }
.ov-docx-p { margin: 0; white-space: pre-wrap; }
.ov-docx-table { border-collapse: collapse; }
.ov-docx-table td { border: 1px solid #999; padding: 3px 5px; vertical-align: top; }
.ov-docx-marker { display: inline-block; min-width: 1.2em; margin-right: .4em; }
.ov-docx-image { display: inline-block; max-width: 100%; }
.ov-docx-image img { max-width: 100%; height: auto; }
.ov-docx-note { font-size: .85em; padding: 4px 0; }
.ov-docx-notes { margin: 20px auto; max-width: 800px; color: #333; }
.ov-search-hit { background: #ffe58f; }
.ov-search-current { background: #ffb300; }
</style></head><body><div class="ov-docx"><div class="ov-docx-scroll"><div class="ov-docx-pages">${body}</div></div></div></body></html>`;
}
