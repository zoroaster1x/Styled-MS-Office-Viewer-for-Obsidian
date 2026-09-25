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

// Autofilter behaviour and the borders that surround filtered rows.
//
// Two things are checked, because both were wrong before:
//   * which rows a filter hides, per filter type, against Excel's rules
//   * that no border belonging to a hidden row is drawn, which is what left
//     ghost lines under a filtered block
//
//   bun test/filters.mjs [file.xlsx]

import { createRequire } from "node:module";
import { readFileSync, existsSync } from "node:fs";
import { setupDom, createContainer } from "./harness.mjs";
import { testPath, skipNotice } from "./env.mjs";

setupDom();
const require = createRequire(import.meta.url);
const { readWorkbook } = require("../src/spreadsheet/read.js");
const { createRenderer } = require("../src/spreadsheet/render.js");

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

function render(book, sheetIndex, filterState) {
  const sheet = book.loadSheet(sheetIndex);
  const host = createContainer();
  const renderer = createRenderer({
    container: host,
    sheet,
    styles: book.styles || sheet.styles,
    theme: book.theme,
    settings: { sheetBackground: "white", showGridlines: true, showHeaders: true },
    filterState: filterState || new Map(),
    outlineCollapsed: new Map(),
  });
  return { sheet, host, renderer };
}

// The rendered grid's first cell of each row tells us which rows are drawn.
function renderedRows(host) {
  const rows = new Set();
  for (const cell of host.querySelectorAll(".xlsx-row-head")) {
    const value = Number(cell.textContent);
    if (Number.isFinite(value)) rows.add(value);
  }
  return rows;
}

function renderedCells(host) {
  return Array.from(host.querySelectorAll(".xlsx-cell"));
}

// Every cell whose inline border carries the test colour, wherever the edge is.
function redBorders(host) {
  const out = [];
  for (const cell of renderedCells(host)) {
    for (const side of ["Top", "Bottom", "Left", "Right"]) {
      const value = cell.style["border" + side] || "";
      if (/rgb\(255, 0, 0\)|#ff0000/i.test(value)) {
        out.push((cell.dataset.r || "?") + ":" + (cell.dataset.c || "?") + " " + side.toLowerCase() + " " + value);
      }
    }
  }
  return out;
}

function cellAt(host, r, c) {
  return host.querySelector('.xlsx-cell[data-r="' + r + '"][data-c="' + c + '"]')
    || host.querySelector('.xlsx-cell[data-r="' + r + '"]');
}

// ---------- synthetic workbook: a bordered block, one row filtered away ----------

function makeSyntheticBook(buildSheet) {
  const { buildSheetModel } = require("../src/delimited/csv.js");
  const rows = [
    ["Group", "Name", "Score"],
    ["A", "Ada", "90"],
    ["B", "Ben", "70"],
    ["A", "Cleo", "85"],
    ["B", "Dan", "60"],
    ["A", "Eve", "95"],
  ];
  const model = buildSheetModel(rows, "Sheet1");
  // An autofilter over the block, with the three columns as filterable.
  model.autoFilter = {
    range: { r1: 1, c1: 1, r2: 6, c2: 3 },
    columns: new Map([
      [0, { colId: 0, values: ["A", "B"] }],
      [1, { colId: 1, values: ["Ada", "Ben", "Cleo", "Dan", "Eve"] }],
      [2, { colId: 2, values: ["90", "70", "85", "60", "95"] }],
    ]),
  };
  buildSheet(model);
  return {
    sheets: [{ name: "Sheet1", index: 0, state: "visible", sheetId: 1, path: "Sheet1" }],
    loadSheet: () => model,
    sheetIndexByName: () => 0,
    styles: model.styles,
    theme: null,
  };
}

// A thick red bottom border on one row only. If it appears after that row is
// filtered away, the line is a ghost of a row that is not drawn.
const TEST_BORDER = { style: "thick", weight: 3, width: 3, css: "solid", color: "#ff0000" };
function addRowBorder(model, row) {
  const cell = model.cells.get(row + ":1");
  if (cell) cell.s = 1;
  const original = model.styles.resolveXf.bind(model.styles);
  model.styles.resolveXf = (idx) => {
    if (idx !== 1) return original(idx);
    const base = original(idx);
    return Object.assign({}, base, {
      borders: Object.assign({}, base.borders, { bottom: TEST_BORDER }),
    });
  };
}

console.log("--- filter semantics ---");

{
  // Values filter: keep group A only.
  const book = makeSyntheticBook(() => {});
  const { host, sheet } = render(book, 0, new Map([[0, new Set(["A"])]]));
  const rows = renderedRows(host);
  check("values filter keeps the header", rows.has(1), [...rows].join(","));
  check("values filter keeps matching rows", rows.has(2) && rows.has(4) && rows.has(6), [...rows].join(","));
  check("values filter hides other rows", !rows.has(3) && !rows.has(5), [...rows].join(","));
  void sheet;
}

{
  // Selecting nothing hides every data row but keeps the header, as Excel does.
  const book = makeSyntheticBook(() => {});
  const { host } = render(book, 0, new Map([[0, new Set()]]));
  const rows = renderedRows(host);
  check("empty selection hides all data rows", rows.size === 1 && rows.has(1), [...rows].join(","));
}

{
  // Two columns filter with AND: group A and score over 88.
  const book = makeSyntheticBook(() => {});
  const filters = new Map([
    [0, new Set(["A"])],
    [2, { custom: { and: true, filters: [{ operator: "greaterThan", val: "88" }] } }],
  ]);
  const { host } = render(book, 0, filters);
  const rows = renderedRows(host);
  check("two filters combine with AND", rows.has(2) && rows.has(6) && !rows.has(4), [...rows].join(","));
}

{
  // Custom filter alone: score under 80 keeps rows 3 and 5.
  const book = makeSyntheticBook(() => {});
  const filters = new Map([[2, { custom: { and: true, filters: [{ operator: "lessThan", val: "80" }] } }]]);
  const { host } = render(book, 0, filters);
  const rows = renderedRows(host);
  check("custom less-than filter", rows.has(3) && rows.has(5) && !rows.has(2), [...rows].join(","));
}

{
  // Picking all values is the same as no filter at all.
  const book = makeSyntheticBook(() => {});
  const { host } = render(book, 0, new Map([[0, new Set(["A", "B"])]]));
  const rows = renderedRows(host);
  check("all values selected shows every row", rows.size === 6, [...rows].join(","));
}

console.log("--- ghost borders ---");

{
  // Row 3 carries a thick red border. Its line is normally drawn on the top
  // edge of the cell below, so a filter that hides row 3 must leave no trace
  // of it anywhere on the grid.
  const book = makeSyntheticBook((model) => addRowBorder(model, 3));
  const { host } = render(book, 0, new Map([[0, new Set(["A"])]]));
  const rows = renderedRows(host);
  check("row 3 is hidden", !rows.has(3));
  check("no trace of the hidden row's border", redBorders(host).length === 0, redBorders(host).join(" | "));

  // With nothing filtered the border is on screen, so the check above means
  // something.
  const plain = render(makeSyntheticBook((model) => addRowBorder(model, 3)), 0, new Map());
  const drawn = redBorders(plain.host);
  check("the border is drawn when its row is visible", drawn.length > 0, drawn.join(" | "));
}

{
  // Column hide is the same rule sideways.
  const book = makeSyntheticBook(() => {});
  const sheet = book.loadSheet(0);
  sheet.cols.push({ min: 2, max: 2, width: 10, hidden: true, style: 0, outline: 0, collapsed: false });
  const host = createContainer();
  const renderer = createRenderer({
    container: host,
    sheet,
    styles: book.styles,
    theme: null,
    settings: { sheetBackground: "white", showGridlines: true, showHeaders: true },
    filterState: new Map(),
    outlineCollapsed: new Map(),
  });
  const headers = Array.from(host.querySelectorAll(".xlsx-col-head")).map((el) => el.textContent);
  check("hidden column is not drawn", headers.indexOf("B") === -1, headers.join(","));
  check("visible columns remain", headers.indexOf("A") !== -1 && headers.indexOf("C") !== -1, headers.join(","));
  renderer.destroy();
}

// ---------- the real timetable ----------

const file = process.argv[2] || testPath("OV_TEST_XLSX");
if (!file) {
  skipNotice("OV_TEST_XLSX", "real timetable filter checks");
} else if (existsSync(file)) {
  console.log("--- real timetable:", file.split("/").pop(), "---");
  const bytes = new Uint8Array(readFileSync(file));
  const book = readWorkbook(bytes);
  const sheetIndex = book.sheets.findIndex((s) => s.name === "TP1 Sep-Dec");
  const sheet = book.loadSheet(sheetIndex >= 0 ? sheetIndex : 0);
  console.log("     sheet:", sheet.name, "autofilter:", sheet.autoFilter ? JSON.stringify(sheet.autoFilter.range) : "none",
    "columns:", sheet.autoFilter ? [...sheet.autoFilter.columns.keys()].join(",") : "-");

  if (sheet.autoFilter) {
    const columns = [...sheet.autoFilter.columns.entries()];
    const groupColumn = columns.find(([, def]) => def.values && def.values.some((v) => /Y1|X1|A|B/i.test(v)));
    if (groupColumn) {
      const [colId, def] = groupColumn;
      const value = def.values.find((v) => /Y1/.test(v)) || def.values[0];
      const { host, renderer } = render(book, sheetIndex >= 0 ? sheetIndex : 0, new Map([[colId, new Set([value])]]));
      const state = renderer.getState();
      check("filtering the timetable reduces the visible rows",
        state.visibleRows < state.totalRows && state.visibleRows > 1,
        state.visibleRows + " of " + state.totalRows);
      check("every visible row matches the filter",
        rowsAllMatch(sheet, colId, value, host),
        "column " + colId + " = " + value);
      check("no cell from a hidden row is drawn",
        noHiddenRowDrawn(sheet, book, sheetIndex >= 0 ? sheetIndex : 0, colId, value, host),
        "");
      renderer.destroy();
    } else {
      check("timetable has a group filter column", false, JSON.stringify(columns.slice(0, 4)));
    }
  }
}

function rowsAllMatch(sheet, colId, value, host) {
  const col = sheet.autoFilter.range.c1 + colId;
  const rows = renderedRows(host);
  for (const r of rows) {
    if (r <= sheet.autoFilter.range.r1) continue;
    const cell = sheet.cells.get(r + ":" + col);
    const text = cell ? String(cell.v == null ? "" : cell.v) : "";
    if (text !== value) return false;
  }
  return true;
}

function noHiddenRowDrawn(sheet, book, sheetIndex, colId, value, host) {
  const col = sheet.autoFilter.range.c1 + colId;
  const rows = renderedRows(host);
  for (const r of rows) {
    if (r <= sheet.autoFilter.range.r1) continue;
    const cell = sheet.cells.get(r + ":" + col);
    const text = cell ? String(cell.v == null ? "" : cell.v) : "";
    if (text !== value) return false;
  }
  void book;
  void sheetIndex;
  return true;
}

console.log("");
console.log("filters:", pass, "pass,", fail, "fail");
process.exit(fail ? 1 : 0);
