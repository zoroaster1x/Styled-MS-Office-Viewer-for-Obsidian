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

// A merged column is stored as D1:D1048576. Taking that span as the used range
// laid out a million rows, and resolving each merged edge walked the whole
// span: a 13 KB timetable took five minutes to open. Both halves are pinned
// here, because either one alone still hangs the view.
//
//   bun test/sheet-dims.mjs

import { createRequire } from "node:module";
import { setupDom, createContainer } from "./harness.mjs";

setupDom();
const require = createRequire(import.meta.url);
const { parseSheet } = require("../src/spreadsheet/sheet.js");
const { buildSheetModel } = require("../src/delimited/csv.js");
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

const SHEET = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:L28"/>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Week</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>One</t></is></c></row>
  </sheetData>
  <mergeCells count="2"><mergeCell ref="D1:D1048576"/><mergeCell ref="A30:B31"/></mergeCells>
</worksheet>`;

const parsed = parseSheet(SHEET, { name: "Sheet1", sharedStrings: [], styles: null, theme: null, rels: new Map() });
check("a whole-column merge does not stretch the used range", parsed.dims.r2 < 500, JSON.stringify(parsed.dims));
check("the used range keeps its own columns", parsed.dims.c2 === 12, JSON.stringify(parsed.dims));
check("a nearby merge still widens the range", parsed.dims.r2 >= 31, JSON.stringify(parsed.dims));

// The renderer half: one merged column on a small sheet must not walk a
// million rows. Before the clamp this took minutes.
const model = buildSheetModel([["a", "b", "c", "d"], ["1", "2", "3", "4"]], "Sheet1");
model.merges = [{ r1: 1, c1: 4, r2: 1048576, c2: 4 }];
model.dims = { r1: 1, c1: 1, r2: 2, c2: 4 };
const started = Date.now();
const host = createContainer();
createRenderer({
  container: host,
  sheet: model,
  styles: model.styles,
  theme: null,
  settings: { sheetBackground: "white", showGridlines: true, showHeaders: true },
  filterState: new Map(),
  outlineCollapsed: new Map(),
});
const elapsed = Date.now() - started;
check("a merged column renders without walking its span", elapsed < 2000, elapsed + " ms");

console.log("");
console.log("sheet-dims: " + pass + " pass, " + fail + " fail");
process.exit(fail ? 1 : 0);
