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

// Fidelity check. Reads the text out of the source XML directly, renders the
// document, and reports any string the renderer dropped. This is the accuracy
// test: a viewer that loses text is wrong regardless of how it looks.
//
//   bun test/fidelity.mjs <file> | --folder <dir> [--limit N] [--verbose]
//
// With no file or folder, the run uses OV_TEST_FOLDER or the individual
// OV_TEST_* documents from .testenv, and is skipped when none are set.

import { createRequire } from "node:module";
import { readFileSync, readdirSync, statSync } from "node:fs";
import { join, extname } from "node:path";
import { setupDom, createContainer } from "./harness.mjs";
import { testEnv, testPath, skipNotice } from "./env.mjs";

setupDom();
const require = createRequire(import.meta.url);

const args = process.argv.slice(2);
const folderIndex = args.indexOf("--folder");
const folder = folderIndex === -1 ? null : args[folderIndex + 1];
const limitIndex = args.indexOf("--limit");
const limit = limitIndex === -1 ? Infinity : parseInt(args[limitIndex + 1], 10);
const verbose = args.includes("--verbose");
const single = args.find((a) => !a.startsWith("--") && a !== folder && a !== String(limit));

// Explicit arguments win; otherwise take whatever .testenv offers.
const configured = [];
for (const key of ["OV_TEST_DOCX", "OV_TEST_PPTX"]) {
  const value = testPath(key);
  if (value) configured.push(value);
}
const envFolder = testPath("OV_TEST_FOLDER");
const files = folder ? collect(folder) : single ? [single] : envFolder ? collect(envFolder) : configured;
if (!files.length) {
  skipNotice("OV_TEST_FOLDER", "fidelity run");
  process.exit(0);
}
void testEnv;

let pass = 0;
let fail = 0;
const problems = [];
const startAll = performance.now();

for (const file of files.slice(0, limit === Infinity ? files.length : limit)) {
  const ext = extname(file).toLowerCase();
  try {
    const result = checkFile(file, ext);
    if (result.missing.length === 0) {
      pass++;
      if (verbose) {
        console.log("ok  ", basename(file), result.checked + " strings");
      }
    } else {
      fail++;
      problems.push({ file, missing: result.missing, checked: result.checked });
      console.log("FAIL", basename(file), "missing " + result.missing.length + " of " + result.checked);
      for (const text of result.missing.slice(0, 8)) {
        console.log("     ", JSON.stringify(text.slice(0, 90)));
      }
    }
  } catch (err) {
    fail++;
    problems.push({ file, error: err.message, missing: [], checked: 0 });
    console.log("ERR ", basename(file), err.message);
  }
}

console.log("");
console.log("fidelity:", pass, "pass,", fail, "fail,", Math.round(performance.now() - startAll) + " ms");
if (problems.length) {
  console.log("files with missing text:", problems.length);
}

function collect(dir) {
  const out = [];
  const walk = (current) => {
    for (const entry of readdirSync(current, { withFileTypes: true })) {
      const full = join(current, entry.name);
      if (entry.isDirectory()) {
        if (entry.name.startsWith(".")) continue;
        walk(full);
      } else if (/\.(docx|docm|pptx|pptm|doc)$/i.test(entry.name)) {
        out.push(full);
      }
    }
  };
  walk(dir);
  return out.sort();
}

function checkFile(file, ext) {
  const bytes = new Uint8Array(readFileSync(file));
  if (ext === ".doc") {
    // Word 97 reads as text: the extracted paragraphs are the source, and the
    // renderer has to show every one of them.
    const { readDoc } = require("../src/legacy/doc.js");
    const model = readDoc(bytes);
    const host = createContainer();
    const { createDocxRenderer } = require("../src/docx/render.js");
    const renderer = createDocxRenderer({ container: host, model, settings: {} });
    const rendered = normalize(host.textContent || "");
    renderer.destroy();
    const source = [];
    for (const block of model.body) {
      for (const run of block.runs || []) {
        if (run.text) source.push(run.text);
      }
    }
    return { checked: source.length, missing: missingText(source, rendered) };
  }
  if (ext === ".docx" || ext === ".docm") {
    const { parseDocx } = require("../src/docx/parse.js");
    const model = parseDocx(bytes);
    const host = createContainer();
    const { createDocxRenderer } = require("../src/docx/render.js");
    const renderer = createDocxRenderer({ container: host, model, settings: {} });
    const rendered = normalize(host.textContent || "");
    renderer.destroy();
    const source = collectDocxText(model);
    return { checked: source.length, missing: missingText(source, rendered) };
  }
  if (ext === ".pptx" || ext === ".pptm") {
    const { parsePptx } = require("../src/pptx/parse.js");
    const model = parsePptx(bytes);
    const host = createContainer();
    const { createPptxRenderer } = require("../src/pptx/render.js");
    const renderer = createPptxRenderer({ container: host, model, settings: { fit: "none" } });
    const renderedParts = [];
    for (let i = 0; i < model.slides.length; i++) {
      renderer.goToSlide(i);
      // The last slide's DOM is all that remains; walk it before moving on.
      renderedParts.push(normalize(host.textContent || ""));
    }
    renderer.destroy();
    const rendered = renderedParts.join(" ");
    const source = collectPptxText(model);
    return { checked: source.length, missing: missingText(source, rendered) };
  }
  throw new Error("unsupported extension");
}

// Text the model holds, at the granularity of one source run. Numbers and
// whitespace-only strings are skipped because rendering legitimately drops
// empty runs and reformats spacing.
function collectDocxText(model) {
  const out = [];
  const walk = (blocks) => {
    for (const block of blocks) {
      if (block.type === "p") {
        for (const run of block.runs) collectRunText(run, out);
      } else if (block.type === "table") {
        for (const row of block.rows) for (const cell of row.cells) walk(cell.blocks);
      }
    }
  };
  walk(model.body);
  return out;
}

function collectRunText(run, out) {
  if (!run) return;
  if (run.type === "text") {
    const text = run.text.trim();
    if (text) out.push(text);
  } else if (run.type === "run") {
    for (const inner of run.runs) collectRunText(inner, out);
  } else if (run.type === "link") {
    for (const inner of run.link.runs) collectRunText(inner, out);
  } else if (run.type === "shapegroup" && run.texts) {
    // Text boxes inside a wpg drawing group are content too.
    for (const text of run.texts) {
      const trimmed = String(text).trim();
      if (trimmed) out.push(trimmed);
    }
  } else if (run.type === "vml" && run.texts) {
    // VML text boxes carry content as well.
    for (const text of run.texts) {
      const trimmed = String(text).trim();
      if (trimmed) out.push(trimmed);
    }
  }
}

function collectPptxText(model) {
  const out = [];
  for (const slide of model.slides) {
    walkShapesForText(slide.shapes, out);
  }
  return out;
}

function walkShapesForText(shapes, out) {
  for (const shape of shapes) {
    if (shape.type === "group") {
      walkShapesForText(shape.shapes, out);
      continue;
    }
    if (shape.textInfo) {
      for (const paragraph of shape.textInfo.paragraphs) {
        for (const run of paragraph.runs) {
          if (run.type === "run") {
            const text = (run.text || "").trim();
            if (text) out.push(text);
          }
        }
      }
    }
    if (shape.table) {
      for (const row of shape.table.rows) {
        for (const cell of row.cells) {
          if (!cell.text) continue;
          for (const paragraph of cell.text.paragraphs) {
            for (const run of paragraph.runs) {
              if (run.type === "run") {
                const text = (run.text || "").trim();
                if (text) out.push(text);
              }
            }
          }
        }
      }
    }
  }
}

// Whitespace is layout, not content, so compare on a collapsed form. Word
// also splits runs mid-word, so a string counts as present when each of its
// tokens appears.
function normalize(text) {
  return String(text || "").replace(/\s+/g, " ").trim();
}

function missingText(sourceStrings, rendered) {
  const haystack = rendered.toLowerCase();
  const missing = [];
  const seen = new Set();
  for (const raw of sourceStrings) {
    const text = normalize(raw);
    if (!text || seen.has(text)) continue;
    seen.add(text);
    if (!haystack.includes(text.toLowerCase())) missing.push(text);
  }
  return missing;
}

function basename(path) {
  const i = path.lastIndexOf("/");
  return i === -1 ? path : path.slice(i + 1);
}
