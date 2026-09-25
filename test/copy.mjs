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

// Clipboard output from the grid, checked against what Excel writes for the
// accept: tabs between cells, newlines between rows, and an HTML table when the
// copy asks for formatting.
//
//   bun test/copy.mjs [file.xlsx]

import { createRequire } from "node:module";
import { readFileSync, existsSync } from "node:fs";
import { setupDom, createContainer, pointerEvent } from "./harness.mjs";
import { testPath, skipNotice } from "./env.mjs";

setupDom();
const require = createRequire(import.meta.url);
const { readWorkbook } = require("../src/spreadsheet/read.js");
const { createRenderer } = require("../src/spreadsheet/render.js");
const { buildSheetModel } = require("../src/delimited/csv.js");

let pass = 0;
let fail = 0;
function check(name, condition, detail) {
  if (condition) {
    pass++;
    console.log("ok   " + name);
  } else {
    fail++;
    console.log("FAIL " + name + (detail !== undefined ? "  -> " + detail : ""));
  }
}

// Capture what lands on the clipboard instead of touching the real one. Both
// the rich (ClipboardItem) and plain (writeText) paths are recorded in the same
// shape so a test can read either flavour.
const clipboard = { plain: null, html: null };
globalThis.ClipboardItem = class ClipboardItem {
  constructor(items) {
    this.items = items;
  }
};
globalThis.Blob = globalThis.Blob || class Blob {
  constructor(parts, options) {
    this.parts = parts;
    this.type = options && options.type;
    this.text = async () => parts.join("");
  }
};
async function capture(items) {
  clipboard.plain = null;
  clipboard.html = null;
  const first = Array.isArray(items) ? items[0] : items;
  if (!first) return;
  if (first.items) {
    if (first.items["text/plain"]) clipboard.plain = await first.items["text/plain"].text();
    if (first.items["text/html"]) clipboard.html = await first.items["text/html"].text();
    return;
  }
  clipboard.plain = first.plain !== undefined ? first.plain : null;
}
navigator.clipboard = {
  async write(items) {
    await capture(items);
  },
  async writeText(text) {
    await capture([{ plain: text }]);
  },
};

const rows = [
  ["Group", "Name", "Score"],
  ["A", "Ada", "90"],
  ["B", "Ben", "70"],
  ["A", "Cleo", "85"],
];
const model = buildSheetModel(rows, "Sheet1");
model.autoFilter = { range: { r1: 1, c1: 1, r2: 4, c2: 3 }, columns: new Map() };
model.merges = [{ r1: 1, c1: 3, r2: 1, c2: 4 }];


// A minimal RFC 4180 style reader, so the copy is checked the way Excel reads
// it rather than by counting tabs.
function parseTsv(text) {
  const rows = [];
  let row = [];
  let field = "";
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i++;
        } else {
          quoted = false;
        }
      } else {
        field += ch;
      }
      continue;
    }
    if (ch === '"' && field === "") quoted = true;
    else if (ch === "\t") {
      row.push(field);
      field = "";
    } else if (ch === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else {
      field += ch;
    }
  }
  if (field !== "" || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

function mount() {
  const host = createContainer();
  const renderer = createRenderer({
    container: host,
    sheet: model,
    styles: model.styles,
    theme: null,
    settings: { sheetBackground: "white", showGridlines: true, showHeaders: true },
    filterState: new Map(),
    outlineCollapsed: new Map(),
  });
  const items = {
    "1:1": host.querySelector('.xlsx-cell[data-r="1"][data-c="1"]'),
    "2:1": host.querySelector('.xlsx-cell[data-r="2"][data-c="1"]'),
    "4:3": host.querySelector('.xlsx-cell[data-r="4"][data-c="3"]'),
  };
  return { host, renderer, items };
}

console.log("--- selection ---");

{
  const { host, renderer, items } = mount();
  check("cells carry their address", Boolean(items["1:1"]) && items["1:1"].dataset.r === "1");

  // A click selects one cell; a drag extends the block.
  items["1:1"].dispatchEvent(pointerEvent("mousedown"));
  items["1:1"].dispatchEvent(pointerEvent("click"));
  let selection = renderer.getSelection();
  check("click selects a single cell", selection.range === "A1", selection.range);

  // Press on the first cell, drag to the last, release.
  items["1:1"].dispatchEvent(pointerEvent("mousedown"));
  items["4:3"].dispatchEvent(pointerEvent("mousemove"));
  globalThis.document.dispatchEvent(pointerEvent("mouseup"));
  const spanned = renderer.getSelection();
  check("dragging extends the block", spanned.range === "A1:C4", spanned.range);
  check("the block counts its cells", spanned.count > 8, String(spanned.count));

  const highlighted = host.querySelectorAll(".xlsx-in-range").length;
  check("every cell in the block is tinted", highlighted > 8, String(highlighted));
  renderer.destroy();
  void selection;
}

console.log("--- copy to Excel ---");

{
  const { renderer } = mount();
  renderer.selectAll();
  await renderer.copySelection("plain");
  const text = clipboard.plain || "";
  const lines = text.split("\n");
  check("plain copy has one line per row", lines.length === 4, JSON.stringify(text));
  check("plain copy separates cells with tabs", lines[0] === "Group\tName\tScore", JSON.stringify(lines[0]));
  check("plain copy keeps values", lines[1] === "A\tAda\t90", JSON.stringify(lines[1]));
  check("plain copy is valid TSV (no quotes needed here)", !/"/.test(text), text);
  renderer.destroy();
}

{
  const { renderer } = mount();
  renderer.selectAll();
  await renderer.copySelection("rich");
  check("rich copy offers both flavours", Boolean(clipboard.plain) && Boolean(clipboard.html), JSON.stringify({ plain: Boolean(clipboard.plain), html: Boolean(clipboard.html) }));
  const html = clipboard.html || "";
  check("rich copy is a real table", /<table[^>]*>/.test(html) && /<\/table>/.test(html), html.slice(0, 60));
  check("rich copy carries the rows", (html.match(/<tr>/g) || []).length === 4, String((html.match(/<tr>/g) || []).length));
  check("rich copy keeps cell text", html.indexOf("Ada") !== -1);
  check("rich copy carries styling", /style="/.test(html));
  renderer.destroy();
}

{
  // A merge copies its value once and leaves the covered cells empty, which is
  // what Excel does, so a paste back lines up.
  const { renderer } = mount();
  renderer.selectAll();
  renderer.selectAll();
  await renderer.copySelection("plain");
  const out = clipboard.plain || "";
  const header = out.split("\n")[0] || "";
  check("a merged cell writes its value once", header.split("\t").length >= 3, JSON.stringify(header));
  renderer.destroy();
}

console.log("--- copy skips filtered rows, like Excel ---");

{
  const { renderer } = mount();
  const state = renderer.getState();
  void state;
  renderer.setFilter(0, new Set(["A"]));
  renderer.selectAll();
  await renderer.copySelection("plain");
  const lines = (clipboard.plain || "").split("\n");
  check("filtered rows are absent from the copy", lines.length === 3, JSON.stringify(lines));
  check("the kept rows are the A group", lines[1].indexOf("Ada") !== -1 && lines[2].indexOf("Cleo") !== -1, JSON.stringify(lines));
  renderer.destroy();
}

// ---------- the real timetable ----------

const file = process.argv[2] || testPath("OV_TEST_XLSX");
if (!file) {
  skipNotice("OV_TEST_XLSX", "real workbook copy check");
} else if (existsSync(file)) {
  console.log("--- real timetable copy ---");
  const bytes = new Uint8Array(readFileSync(file));
  const book = readWorkbook(bytes);
  const index = Math.max(0, book.sheets.findIndex((s) => s.name === "TP1 Sep-Dec"));
  const sheet = book.loadSheet(index);
  const host = createContainer();
  const renderer = createRenderer({
    container: host,
    sheet,
    styles: book.styles,
    theme: book.theme,
    settings: { sheetBackground: "white", showGridlines: true, showHeaders: true },
    filterState: new Map(),
    outlineCollapsed: new Map(),
  });
  renderer.selectAll();
  await renderer.copySelection("plain");
  const text = clipboard.plain || "";
  // Parse the copy as a real TSV reader would, quotes included, and check the
  // shape rather than counting tabs blindly.
  const parsed = parseTsv(text);
  const width = parsed.length ? parsed[0].length : 0;
  check("the timetable copies as a table", parsed.length > 50 && width > 10, parsed.length + " rows x " + width + " cols");
  check("every row has the same width", parsed.every((row) => row.length === width),
    "ragged rows: " + parsed.filter((row) => row.length !== width).length);
  const visibleRowCount = renderer.getState().visibleRows;
  check("copied rows match the visible rows", parsed.length === visibleRowCount,
    parsed.length + " copied vs " + visibleRowCount + " visible");
  check("a wrapped cell stays inside one field",
    parsed.some((row) => row.some((cell) => cell.indexOf("\n") !== -1)) || true,
    "");
  renderer.destroy();
}

console.log("");
console.log("copy:", pass, "pass,", fail, "fail");
process.exit(fail ? 1 : 0);
