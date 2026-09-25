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

// Column and row sizing: dragging a header edge, auto fitting, the remembered
// sizes, and the reader-set minimums.
//
//   bun test/sizing.mjs [file.xlsx]

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

const rows = [
  ["Short", "A much longer value that will not fit a default column width at all"],
  ["Also short", "Another long value to measure the auto fit against"],
];
const model = buildSheetModel(rows, "Sizes");

function mount(settings, overrides) {
  const host = createContainer();
  const renderer = createRenderer({
    container: host,
    sheet: model,
    styles: model.styles,
    theme: null,
    settings: Object.assign({ sheetBackground: "white", showGridlines: true, showHeaders: true }, settings || {}),
    sizeOverrides: overrides,
    filterState: new Map(),
    outlineCollapsed: new Map(),
  });
  return { host, renderer };
}

function templateCols(host) {
  const grid = host.querySelector(".xlsx-sheet-grid");
  return (grid.style.gridTemplateColumns || "").split(/\s+/).filter(Boolean);
}

console.log("--- dragging a column edge ---");

{
  const { host, renderer } = mount();
  // The first handle is column A, which the template lists at index 1.
  const before = templateCols(host)[1];
  const handle = host.querySelector(".xlsx-col-resize");
  check("a column handle exists", Boolean(handle));
  if (handle) {
    handle.dispatchEvent(pointerEvent("mousedown", { clientX: 500 }));
    globalThis.document.dispatchEvent(pointerEvent("mousemove", { clientX: 560 }));
    globalThis.document.dispatchEvent(pointerEvent("mouseup", { clientX: 560 }));
    const after = templateCols(host)[1];
    check("the column width changed while dragging", before !== after, before + " -> " + after);
    const width = parseInt(after, 10);
    check("the width matches the drag distance", Math.abs(width - (parseInt(before, 10) + 60)) <= 2, after);
    const sizes = renderer.getSizes();
    check("the width is recorded for this sheet", sizes.cols[1] === width, JSON.stringify(sizes.cols));
  }
  renderer.destroy();
}

console.log("--- the size is remembered ---");

{
  const { host, renderer } = mount();
  const handle = host.querySelector(".xlsx-col-resize");
  handle.dispatchEvent(pointerEvent("mousedown", { clientX: 500 }));
  globalThis.document.dispatchEvent(pointerEvent("mousemove", { clientX: 580 }));
  globalThis.document.dispatchEvent(pointerEvent("mouseup", { clientX: 580 }));
  const sizes = renderer.getSizes();
  const width = sizes.cols[1];
  check("the drag recorded a width", Number.isFinite(width) && width > 0, JSON.stringify(sizes));
  renderer.destroy();

  // A new renderer given those sizes starts with them, which is what happens
  // when the file is reopened.
  const { host: reopened, renderer: second } = mount(null, sizes);
  check("a remembered width is applied on open", templateCols(reopened)[1] === width + "px",
    "expected " + width + "px, template " + JSON.stringify(templateCols(reopened)) + ", sizes " + JSON.stringify(sizes));
  check("the renderer reports the same sizes", second.getSizes().cols[1] === width, JSON.stringify(second.getSizes()));
  second.destroy();
}

console.log("--- dragging a row edge ---");

{
  const { host, renderer } = mount();
  const grid = host.querySelector(".xlsx-sheet-grid");
  const before = (grid.style.gridTemplateRows || "").split(/\s+/)[1];
  const handle = host.querySelector(".xlsx-row-resize");
  check("a row handle exists", Boolean(handle));
  if (handle) {
    handle.dispatchEvent(pointerEvent("mousedown", { clientY: 100 }));
    globalThis.document.dispatchEvent(pointerEvent("mousemove", { clientY: 140 }));
    globalThis.document.dispatchEvent(pointerEvent("mouseup", { clientY: 140 }));
    const after = (grid.style.gridTemplateRows || "").split(/\s+/)[1];
    check("the row height changed", before !== after, before + " -> " + after);
    check("the height is recorded", renderer.getSizes().rows[1] >= 40, JSON.stringify(renderer.getSizes().rows));
  }
  renderer.destroy();
}

console.log("--- auto fit on double click ---");

{
  const { host, renderer } = mount();
  // Column B holds the long value, so auto fitting it must widen the column.
  const before = parseInt(templateCols(host)[2], 10);
  renderer.autoFitColumn(2);
  const after = parseInt(templateCols(host)[2], 10);
  check("auto fit widens the column for the long value", after > before, before + " -> " + after);
  check("auto fit stops at a sane maximum", after <= 900, String(after));
  check("auto fit is recorded", renderer.getSizes().cols[2] === after, JSON.stringify(renderer.getSizes().cols));

  // And the double click on a handle does the same as the API.
  const handle = host.querySelector(".xlsx-col-resize");
  if (handle) {
    renderer.resetSizes();
    handle.dispatchEvent(pointerEvent("dblclick"));
    check("double clicking the handle auto fits column A", parseInt(templateCols(host)[1], 10) > 0,
      templateCols(host)[1]);
  } else {
    check("a column handle exists for auto fit", false);
  }
  renderer.destroy();
}

console.log("--- resetting ---");

{
  const { host, renderer } = mount();
  renderer.autoFitColumn(2);
  const wide = parseInt(templateCols(host)[2], 10);
  renderer.resetSizes();
  const normal = parseInt(templateCols(host)[2], 10);
  check("reset returns the column to the file width", normal < wide, wide + " -> " + normal);
  check("reset clears the record", Object.keys(renderer.getSizes().cols).length === 0, JSON.stringify(renderer.getSizes()));
  renderer.destroy();
}

console.log("--- the minimum width setting ---");

{
  const { host, renderer } = mount({ minColumnWidth: 220 });
  const width = parseInt(templateCols(host)[2], 10);
  check("the minimum width is applied", width >= 220, String(width));
  const rowHeight = parseFloat((host.querySelector(".xlsx-sheet-grid").style.gridTemplateRows || "").split(/\s+/)[1]);
  check("rows keep the file height when no minimum is set", rowHeight > 0, String(rowHeight));
  renderer.destroy();

  const { host: host2, renderer: second } = mount({ minRowHeight: 44 });
  const rowHeight2 = parseFloat((host2.querySelector(".xlsx-sheet-grid").style.gridTemplateRows || "").split(/\s+/)[1]);
  check("the minimum row height is applied", rowHeight2 >= 44, String(rowHeight2));
  second.destroy();
}

console.log("--- a dragged size wins over the minimum ---");

{
  const { host, renderer } = mount({ minColumnWidth: 300 });
  check("the minimum applies before any fit", parseInt(templateCols(host)[1], 10) >= 300, templateCols(host)[1]);
  renderer.autoFitColumn(1);
  const width = parseInt(templateCols(host)[1], 10);
  check("a narrower auto fit overrides the minimum for that column", width < 300, String(width));
  renderer.destroy();
}

// ---------- the real timetable ----------

const file = process.argv[2] || testPath("OV_TEST_XLSX");
if (!file) {
  skipNotice("OV_TEST_XLSX", "real timetable sizing checks");
} else if (existsSync(file)) {
  console.log("--- real timetable ---");
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
    settings: { sheetBackground: "white", showGridlines: true, showHeaders: true, minColumnWidth: 0 },
    filterState: new Map(),
    outlineCollapsed: new Map(),
  });
  const handles = host.querySelectorAll(".xlsx-col-resize").length;
  check("every visible column has a handle", handles > 10, String(handles));
  const wide = host.querySelectorAll(".xlsx-row-resize").length;
  check("visible rows have handles", wide > 10, String(wide));
  const templateCount = (host.querySelector(".xlsx-sheet-grid").style.gridTemplateColumns || "").split(/\s+/).length;
  check("the template still matches the columns", templateCount >= 24, String(templateCount));
  renderer.destroy();
}

console.log("");
console.log("sizing:", pass, "pass,", fail, "fail");
process.exit(fail ? 1 : 0);
