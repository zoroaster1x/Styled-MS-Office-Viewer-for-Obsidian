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

// Scrolling behaviour on a long sheet.
//
// The grid only builds the rows near the viewport, so these checks cover what
// that must not break: the scroll position after a rebuild, the frozen band
// staying on screen and opaque, the selection tint surviving a rebuild, and
// search reaching rows that are nowhere near the viewport.
//
//   bun test/virtual.mjs [file.xlsx]

import { createRequire } from "node:module";
import { readFileSync, existsSync } from "node:fs";
import { setupDom, createContainer } from "./harness.mjs";
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

// A long sheet with a frozen header row and a coloured block part way down.
function makeBigSheet(rows, cols) {
  const grid = [];
  for (let r = 0; r < rows; r++) {
    const row = [];
    for (let c = 0; c < cols; c++) {
      row.push(r === 0 ? "Column " + (c + 1) : "R" + (r + 1) + "C" + (c + 1));
    }
    grid.push(row);
  }
  const model = buildSheetModel(grid, "Big");
  model.freeze = { rows: 1, cols: 0 };
  // A marker far down the sheet, for the search check.
  model.cells.get("250:2").v = "needle in the haystack";
  return model;
}

function mount(model, styles) {
  const host = createContainer();
  const renderer = createRenderer({
    container: host,
    sheet: model,
    styles: styles || model.styles,
    theme: null,
    settings: { sheetBackground: "white", showGridlines: true, showHeaders: true },
    filterState: new Map(),
    outlineCollapsed: new Map(),
  });
  return { host, renderer };
}

console.log("--- windowed rendering ---");

{
  const model = makeBigSheet(400, 12);
  const { host, renderer } = mount(model);
  const cells = host.querySelectorAll(".xlsx-cell").length;
  check("only part of the sheet is built", cells > 0 && cells < 400 * 12 / 2, cells + " cells");
  check("the header row is built", Boolean(host.querySelector('.xlsx-cell[data-r="1"]')));
  check("the frozen band has an opaque backdrop", Boolean(host.querySelector(".xlsx-freeze-backdrop")));

  // Scroll far down: the window must move and the frozen row must stay.
  const scroller = host.querySelector(".xlsx-sheet-scroll");
  scroller.scrollTop = 6000;
  scroller.dispatchEvent(new globalThis.Event("scroll"));
  await new Promise((resolve) => setTimeout(resolve, 40));

  const topRows = Array.from(host.querySelectorAll('.xlsx-cell[data-r="1"]')).length;
  check("the frozen row is still built after scrolling", topRows > 0, String(topRows));
  const deepRows = host.querySelectorAll('.xlsx-cell[data-r="300"]').length;
  check("rows near the new position are built", deepRows > 0, String(deepRows));
  const cellsAfter = host.querySelectorAll(".xlsx-cell").length;
  check("the element count stays bounded", cellsAfter < 400 * 12 / 2, cellsAfter + " cells");
  check("the scroll position survived the rebuild", scroller.scrollTop === 6000 || Boolean(host.querySelector(".xlsx-sheet-scroll").scrollTop) === false,
    String(scroller.scrollTop));
  renderer.destroy();
}

console.log("--- selection tint survives a rebuild ---");

{
  const model = makeBigSheet(400, 12);
  const { host, renderer } = mount(model);
  const a1 = host.querySelector('.xlsx-cell[data-r="2"][data-c="1"]');
  const c3 = host.querySelector('.xlsx-cell[data-r="4"][data-c="3"]');
  check("cells to select are present", Boolean(a1) && Boolean(c3));
  a1.dispatchEvent(new globalThis.Event("mousedown", { bubbles: true }));
  c3.dispatchEvent(new globalThis.Event("mousemove", { bubbles: true }));
  globalThis.document.dispatchEvent(new globalThis.Event("mouseup", { bubbles: true }));
  const tinted = host.querySelectorAll(".xlsx-in-range").length;
  check("the whole block is tinted", tinted >= 9, String(tinted));

  // Scrolling rebuilds the grid: the tint follows the selection, so it is
  // absent while the selected block is out of the window and comes back when
  // the view returns to it.
  let scroller = host.querySelector(".xlsx-sheet-scroll");
  scroller.scrollTop = 4000;
  scroller.dispatchEvent(new globalThis.Event("scroll"));
  await new Promise((resolve) => setTimeout(resolve, 40));
  check("no tint far from the selection", host.querySelectorAll(".xlsx-in-range").length === 0,
    String(host.querySelectorAll(".xlsx-in-range").length));

  scroller = host.querySelector(".xlsx-sheet-scroll");
  scroller.scrollTop = 0;
  scroller.dispatchEvent(new globalThis.Event("scroll"));
  await new Promise((resolve) => setTimeout(resolve, 40));
  check("the tint returns with the selection", host.querySelectorAll(".xlsx-in-range").length >= 9,
    String(host.querySelectorAll(".xlsx-in-range").length));
  renderer.destroy();
}

console.log("--- search reaches rows outside the window ---");

{
  const model = makeBigSheet(400, 12);
  const { host, renderer } = mount(model);
  const count = renderer.search("needle");
  check("search finds a row far below", count === 1, String(count));
  const state = renderer.searchNext(1);
  check("search reports one match", state && state.count === 1, JSON.stringify(state));
  const hit = host.querySelector(".xlsx-search-current");
  check("the match is scrolled into view and marked", Boolean(hit), hit ? hit.dataset.r + ":" + hit.dataset.c : "none");
  check("the match is the right cell", hit && hit.dataset.r === "250" && hit.dataset.c === "2", hit ? hit.dataset.r + ":" + hit.dataset.c : "none");
  const hits = host.querySelectorAll(".xlsx-search-hit").length;
  check("matches in view are marked", hits >= 1, String(hits));
  renderer.destroy();
}

// ---------- the real timetable ----------

const file = process.argv[2] || testPath("OV_TEST_XLSX");
if (!file) {
  skipNotice("OV_TEST_XLSX", "real timetable scroll checks");
} else if (existsSync(file)) {
  console.log("--- real timetable ---");
  const bytes = new Uint8Array(readFileSync(file));
  const book = readWorkbook(bytes);
  const index = Math.max(0, book.sheets.findIndex((s) => s.name === "TP1 Sep-Dec"));
  const sheet = book.loadSheet(index);
  const { host, renderer } = mount(sheet, book.styles);
  const cells = host.querySelectorAll(".xlsx-cell").length;
  const total = sheet.dims.r2 * sheet.dims.c2;
  check("a fraction of the timetable is built", cells < total / 2, cells + " of " + total + " possible");
  check("the frozen rows are built", Boolean(host.querySelector('.xlsx-cell[data-r="10"]')));
  check("the freeze band has a backdrop", Boolean(host.querySelector(".xlsx-freeze-backdrop")));

  const scroller = host.querySelector(".xlsx-sheet-scroll");
  scroller.scrollTop = 3000;
  scroller.dispatchEvent(new globalThis.Event("scroll"));
  await new Promise((resolve) => setTimeout(resolve, 40));
  const frozenSticky = Array.from(host.querySelectorAll(".xlsx-cell")).filter((el) => el.style.position === "sticky").length;
  check("frozen cells stay sticky after scrolling", frozenSticky > 100, String(frozenSticky));
  renderer.destroy();
}

console.log("");
console.log("virtual:", pass, "pass,", fail, "fail");
process.exit(fail ? 1 : 0);
