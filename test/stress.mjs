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

// Stress harness: parses and renders every Office document under a folder and
// reports timing, structure counts and anything that throws. This is how the
// parser is validated against real coursework without opening Obsidian.
//
//   bun test/stress.mjs <folder> [--limit 20] [--render] [--html dir]
//
// With no folder argument the run uses OV_TEST_FOLDER from .testenv, and is
// skipped when that is not set.

import { createRequire } from "node:module";
import { readdirSync, statSync, mkdirSync, writeFileSync } from "node:fs";
import { join, extname, basename } from "node:path";
import { setupDom, createContainer } from "./harness.mjs";
import { testPath, skipNotice } from "./env.mjs";

setupDom();
const require = createRequire(import.meta.url);

const args = process.argv.slice(2);
const root = args.find((a) => !a.startsWith("--")) || testPath("OV_TEST_FOLDER");
if (!root) {
  skipNotice("OV_TEST_FOLDER", "stress run");
  process.exit(0);
}
const limitArg = args.indexOf("--limit");
const limit = limitArg === -1 ? Infinity : parseInt(args[limitArg + 1], 10);
const doRender = args.includes("--render");
const htmlArg = args.indexOf("--html");
const htmlDir = htmlArg === -1 ? null : args[htmlArg + 1];
if (htmlDir) mkdirSync(htmlDir, { recursive: true });

const OFFICE = new Set([".docx", ".docm", ".pptx", ".pptm", ".xlsx", ".xlsm", ".doc"]);

function walk(dir, out) {
  let entries;
  try {
    entries = readdirSync(dir, { withFileTypes: true });
  } catch (err) {
    return out;
  }
  for (const entry of entries) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) {
      if (entry.name.startsWith(".") || entry.name === "node_modules") continue;
      walk(full, out);
    } else if (entry.isFile() && OFFICE.has(extname(entry.name).toLowerCase())) {
      out.push(full);
    }
  }
  return out;
}

const files = walk(root, []).sort();
const selected = files.slice(0, limit === Infinity ? files.length : limit);
console.log("folder:", root);
console.log("office files found:", files.length, "testing:", selected.length);
console.log("");

const results = { ok: 0, fail: 0, byExt: new Map(), totalMs: 0 };
const failures = [];
const slowest = [];

for (const file of selected) {
  const ext = extname(file).toLowerCase();
  const entry = { file, ms: 0, parseMs: 0, renderMs: 0, stats: null, html: 0, error: null };
  try {
    const { readFileSync } = require("node:fs");
    const bytes = new Uint8Array(readFileSync(file));
    let model;
    let t0 = performance.now();
    if (ext === ".doc") {
      model = require("../src/legacy/doc.js").readDoc(bytes);
    } else if (ext === ".docx" || ext === ".docm") {
      model = require("../src/docx/parse.js").parseDocx(bytes);
    } else if (ext === ".pptx" || ext === ".pptm") {
      model = require("../src/pptx/parse.js").parsePptx(bytes);
    } else {
      model = require("../src/spreadsheet/read.js").readWorkbook(bytes);
    }
    entry.parseMs = performance.now() - t0;
    entry.stats = summarize(model, ext);
    if (doRender) {
      t0 = performance.now();
      const host = createContainer();
      if (ext === ".docx" || ext === ".docm" || ext === ".doc") {
        const { createDocxRenderer } = require("../src/docx/render.js");
        const renderer = createDocxRenderer({ container: host, model, settings: {} });
        entry.html = host.innerHTML.length;
        if (htmlDir) writeFileSync(join(htmlDir, sanitize(basename(file)) + ".html"), host.innerHTML);
        renderer.destroy();
      } else if (ext === ".pptx" || ext === ".pptm") {
        const { createPptxRenderer } = require("../src/pptx/render.js");
        const renderer = createPptxRenderer({
          container: host,
          model,
          settings: { fit: "none" },
          onSlideChange: () => {},
        });
        // Draw a few slides so picture media and text layout are exercised.
        const total = model.slides.length;
        for (const index of [0, Math.floor(total / 2), total - 1]) {
          if (index >= 0 && index < total) renderer.goToSlide(index);
        }
        entry.html = host.innerHTML.length;
        if (htmlDir) writeFileSync(join(htmlDir, sanitize(basename(file)) + ".html"), host.innerHTML);
        renderer.destroy();
      }
      entry.renderMs = performance.now() - t0;
    }
    entry.ms = entry.parseMs + entry.renderMs;
    results.ok++;
    results.totalMs += entry.ms;
    const bucket = results.byExt.get(ext) || { count: 0, ms: 0, blocks: 0, images: 0 };
    bucket.count++;
    bucket.ms += entry.ms;
    bucket.blocks += entry.stats.blocks;
    bucket.images += entry.stats.images;
    results.byExt.set(ext, bucket);
    slowest.push(entry);
  } catch (err) {
    entry.error = err && err.stack ? err.stack.split("\n").slice(0, 3).join(" | ") : String(err);
    results.fail++;
    failures.push(entry);
  }
}

slowest.sort((a, b) => b.ms - a.ms);
console.log("--- summary ---");
console.log("ok:", results.ok, "fail:", results.fail, "total:", Math.round(results.totalMs) + " ms");
for (const [ext, bucket] of results.byExt) {
  console.log(
    ext.padEnd(6),
    bucket.count + " files",
    Math.round(bucket.ms) + " ms total",
    (bucket.ms / Math.max(1, bucket.count)).toFixed(1) + " ms avg",
    bucket.blocks + " blocks",
    bucket.images + " images",
  );
}
console.log("");
console.log("--- slowest ---");
for (const entry of slowest.slice(0, 12)) {
  console.log(
    entry.ms.toFixed(0).padStart(6) + " ms",
    "parse " + entry.parseMs.toFixed(0).padStart(5) + " render " + entry.renderMs.toFixed(0).padStart(5),
    (entry.html ? "html " + entry.html : "").padEnd(14),
    basename(entry.file).slice(0, 70),
  );
}
if (failures.length) {
  console.log("");
  console.log("--- failures ---");
  for (const entry of failures) {
    console.log("!!", entry.file);
    console.log("   ", entry.error);
  }
}

function summarize(model, ext) {
  if (ext === ".doc") {
    // Word 97 is read as text: paragraphs and their words.
    let blocks = 0;
    let paragraphs = 0;
    let characters = 0;
    for (const block of model.body || []) {
      blocks++;
      if (block.type === "p" && block.runs && block.runs.length) {
        paragraphs++;
        for (const run of block.runs) characters += (run.text || "").length;
      }
    }
    return { blocks, paragraphs, characters, images: 0 };
  }
  if (ext === ".docx" || ext === ".docm") {
    let blocks = 0, images = 0, tables = 0, paragraphs = 0, lists = 0;
    const walk = (list) => {
      for (const block of list || []) {
        blocks++;
        if (block.type === "p") {
          paragraphs++;
          if (block.numbering) lists++;
          for (const run of block.runs || []) countRun(run);
        } else if (block.type === "table") {
          tables++;
          for (const row of block.rows) for (const cell of row.cells) walk(cell.blocks);
        }
      }
    };
    const countRun = (run) => {
      if (!run) return;
      if (run.type === "image") images++;
      else if (run.type === "run") for (const inner of run.runs) countRun(inner);
      else if (run.type === "link") for (const inner of run.link.runs) countRun(inner);
    };
    walk(model.body);
    return { blocks, images, tables, paragraphs, lists };
  }
  if (ext === ".pptx" || ext === ".pptm") {
    let shapes = 0, images = 0, text = 0, tables = 0;
    const count = (list) => {
      for (const shape of list) {
        shapes++;
        if (shape.type === "picture") images++;
        if (shape.type === "table") tables++;
        if (shape.text) text += shape.text.length;
        if (shape.type === "group") count(shape.shapes);
      }
    };
    for (const slide of model.slides) count(slide.shapes);
    return { blocks: shapes, images, slides: model.slides.length, text, tables };
  }
  return { blocks: model.sheets.length, images: 0, sheets: model.sheets.length };
}

function sanitize(name) {
  return name.replace(/[^\w.-]+/g, "_").slice(0, 120);
}
